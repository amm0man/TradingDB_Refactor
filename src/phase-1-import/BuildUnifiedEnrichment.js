/**
 * BuildUnifiedEnrichment.js
 *
 * Phase 1 – Group A: fee enrichment & timestamp matching
 *
 * Sole responsibility:
 *   Match TosTrades legs to the correct TosTop TRD fee/amount rows.
 *   Consumes items from the in-memory TosTop queues so each TRD is used once.
 *
 * Why a factory (createTopTradeEnrichmentHelpers):
 *   These helpers are not standalone. They pull items out of two maps that
 *   buildUnifiedImportV3() builds in section 5:
 *     - topTradeQueueByKey
 *     - topTradeQueueByDateTime
 *   Passing those maps in keeps the matching logic identical to the nested
 *   version, without making the queues global.
 *
 * Called by:
 *   buildUnifiedImportV3() after the two queues exist
 *
 * Shared helpers used from Helpers.js (already global):
 *   toNum, normalizeDate, normalizeTime, normalizeTimeHHmmss, toDateObject
 *
 * Still passed in from buildUnifiedImportV3 (not moved yet):
 *   toStr, makeTradeMatchKey
 *
 * Related files:
 *   - BuildUnifiedImportV3.js      (orchestrator / remaining nested groups)
 *   - TosSchwabImportPipeline.js   (fills TosTrades / TosTop)
 *   - ImportIssues.js
 */

/**
 * createTopTradeEnrichmentHelpers(opts)
 *
 * opts.topTradeQueueByKey      – exact-second index (mutated as rows are consumed)
 * opts.topTradeQueueByDateTime – minute-bucket index (mutated as rows are consumed)
 * opts.toStr                   – nested string helper from buildUnifiedImportV3
 * opts.makeTradeMatchKey       – nested exact-key builder from buildUnifiedImportV3
 *
 * Returns the same five function names that used to be nested in
 * buildUnifiedImportV3. Call-site argument lists are unchanged.
 */
function createTopTradeEnrichmentHelpers(opts) {
  const topTradeQueueByKey = opts.topTradeQueueByKey;
  const topTradeQueueByDateTime = opts.topTradeQueueByDateTime;
  const toStr = opts.toStr;
  const makeTradeMatchKey = opts.makeTradeMatchKey;

  function previewTopCandidates(list, limit) {
    const arr = Array.isArray(list) ? list : [];
    const n = Math.min(arr.length, limit || 6);
    const out = [];
    for (let i = 0; i < n; i++) {
      const it = arr[i];
      out.push({
        topSym: it.topSym,
        topAbsQty: it.topAbsQty,
        topPrice: it.topPrice,
        topExactKey: it.topExactKey,
        topDesc: String(it.topDesc || "").substring(0, 120),
      });
    }
    return out;
  }

  function removeFromExactIndex(item) {
    if (!item || !item.topExactKey) return;
    const b = topTradeQueueByKey[item.topExactKey];
    if (!b || !b.length) return;
    const j = b.indexOf(item);
    if (j >= 0) b.splice(j, 1);
  }

  /**
   * Combine TosTop partial fills into one enrichment pull.
   *
   * Used when no single TosTop TRD has the full TosTrades qty, but 2+ unused
   * TRDs for the same underlying + same price add up exactly to that qty.
   * Searches the trade minute and the minutes immediately before/after
   * (fills often print a few seconds apart and can cross :59 -> :00).
   *
   * Example: TosTrades SELL -76 UUUU at 13:15:05
   *          TosTop     SOLD -10 at 13:14:56 + SOLD -66 at 13:15:05
   */
  function pullPartialFillEnrichment_(
    account,
    dateIso,
    tradeTs,
    minuteKey,
    sym,
    qtyAbs,
    matchPrice,
    debug,
  ) {
    const wantQty = Number(qtyAbs);
    if (!isFinite(wantQty) || wantQty <= 0) return null;

    const nearbyKeys = [minuteKey];
    if (tradeTs instanceof Date && !isNaN(tradeTs.getTime())) {
      const minusTs = new Date(tradeTs.getTime() - 60 * 1000);
      const plusTs = new Date(tradeTs.getTime() + 60 * 1000);
      nearbyKeys.push(
        [account, normalizeDate(minusTs), normalizeTime(minusTs)].join("|"),
      );
      nearbyKeys.push(
        [account, normalizeDate(plusTs), normalizeTime(plusTs)].join("|"),
      );
    }

    const pieces = [];
    for (let k = 0; k < nearbyKeys.length; k++) {
      const dtKey = nearbyKeys[k];
      const bucket = topTradeQueueByDateTime[dtKey] || [];
      for (let i = 0; i < bucket.length; i++) {
        const it = bucket[i];
        if (!it) continue;
        if (
          String(it.topSym || "")
            .trim()
            .toUpperCase() !== sym
        )
          continue;

        const pieceQty = Number(it.topAbsQty);
        if (!isFinite(pieceQty) || pieceQty <= 0) continue;

        const a = toNum(it.topPrice);
        const b = toNum(matchPrice);
        if (!isNaN(a) && !isNaN(b) && Math.abs(a - b) > 0.0001) continue;

        let dist = 999999999;
        if (
          tradeTs instanceof Date &&
          !isNaN(tradeTs.getTime()) &&
          it.topTs instanceof Date &&
          !isNaN(it.topTs.getTime())
        ) {
          dist = Math.abs(it.topTs.getTime() - tradeTs.getTime());
        }

        pieces.push({
          it: it,
          bucket: bucket,
          idx: i,
          dtKey: dtKey,
          pieceQty: pieceQty,
          dist: dist,
        });
      }
    }

    debug.counts.symbolCandidatesNearby = pieces.length;

    if (!pieces.length) return null;

    // Closest prints first, then earlier-in-bucket as a tie-break.
    pieces.sort(function (a, b) {
      if (a.dist !== b.dist) return a.dist - b.dist;
      if (a.dtKey !== b.dtKey) return a.idx - b.idx;
      return a.idx - b.idx;
    });

    const picked = [];
    let sumQty = 0;
    for (let i = 0; i < pieces.length; i++) {
      const piece = pieces[i];
      if (sumQty + piece.pieceQty > wantQty) continue;
      picked.push(piece);
      sumQty += piece.pieceQty;
      if (sumQty === wantQty) break;
    }

    debug.counts.partialFillQtySum = sumQty;
    debug.counts.partialFillPicked = picked.length;

    if (sumQty !== wantQty || !picked.length) return null;

    let sumMiscFees = 0;
    let sumFeesComm = 0;
    let sumAmount = 0;
    let sawMiscFees = false;
    let sawFeesComm = false;
    let sawAmount = false;

    // Remove last-in-bucket first so earlier indexes stay valid.
    picked.sort(function (a, b) {
      if (a.dtKey !== b.dtKey) return a.dtKey < b.dtKey ? -1 : 1;
      return b.idx - a.idx;
    });

    for (let i = 0; i < picked.length; i++) {
      const piece = picked[i];
      const removed = piece.bucket.splice(piece.idx, 1)[0];
      if (!removed) continue;
      removeFromExactIndex(removed);

      const mf = toNum(removed.miscFees);
      const fc = toNum(removed.feesComm);
      const am = toNum(removed.amount);
      if (!isNaN(mf)) {
        sumMiscFees += mf;
        sawMiscFees = true;
      }
      if (!isNaN(fc)) {
        sumFeesComm += fc;
        sawFeesComm = true;
      }
      if (!isNaN(am)) {
        sumAmount += am;
        sawAmount = true;
      }
    }

    debug.why =
      "Matched partial fills: same symbol + same price, qty pieces sum to TosTrades qty.";

    return {
      item: {
        miscFees: sawMiscFees ? sumMiscFees : "",
        feesComm: sawFeesComm ? sumFeesComm : "",
        amount: sawAmount ? sumAmount : "",
        matchedBy: "partialFillQtySum",
        matchedKey: minuteKey,
        pickedCount: picked.length,
      },
      debug: debug,
    };
  }

  function pullTopTradeEnrichment(
    Account,
    tradeTs,
    sym,
    qtyAbs,
    matchPrice,
    expectedFillCount,
  ) {
    expectedFillCount = Number(expectedFillCount || 1);

    const al = toStr(Account).trim().toUpperCase();
    const s = toStr(sym).trim().toUpperCase();
    const q = Number(qtyAbs);

    const dateIso = normalizeDate(tradeTs);
    const timeHHmm = normalizeTime(tradeTs);
    const timeHHmmss = normalizeTimeHHmmss(tradeTs);

    const minuteKey = [al, dateIso, timeHHmm].join("|");
    const exactKey = makeTradeMatchKey(
      al,
      dateIso,
      timeHHmmss,
      s,
      q,
      matchPrice,
    );

    const debug = {
      minuteKey: minuteKey,
      exactKey: exactKey,
      requested: {
        Account: al,
        dateIso: dateIso,
        timeHHmm: timeHHmm,
        timeHHmmss: timeHHmmss,
        sym: s,
        qtyAbs: q,
        matchPrice: matchPrice,
      },
      counts: {},
      why: "",
    };

    // A) Exact second key fast path
    const exactBucket = topTradeQueueByKey[exactKey];
    debug.counts.exactBucket = exactBucket ? exactBucket.length : 0;

    if (exactBucket && exactBucket.length) {
      const item = exactBucket.shift();
      const mb = topTradeQueueByDateTime[item.topMinuteKey];
      if (mb && mb.length) {
        const j = mb.indexOf(item);
        if (j >= 0) mb.splice(j, 1);
      }
      item.matchedBy = "exactSecondKey";
      item.matchedKey = exactKey;
      debug.why = "Matched exact second key.";
      return { item: item, debug: debug };
    }

    // B) Minute bucket closest-in-time (with +/-1 minute fallback)
    let dtBucket = topTradeQueueByDateTime[minuteKey] || [];
    debug.counts.minuteBucket = dtBucket.length;

    let minuteKeyUsed = minuteKey;

    if (!dtBucket.length) {
      const minusTs = new Date(tradeTs.getTime() - 60 * 1000);
      const plusTs = new Date(tradeTs.getTime() + 60 * 1000);

      const minuteKeyMinus = [
        al,
        normalizeDate(minusTs),
        normalizeTime(minusTs),
      ].join("|");
      const minuteKeyPlus = [
        al,
        normalizeDate(plusTs),
        normalizeTime(plusTs),
      ].join("|");

      const bucketMinus = topTradeQueueByDateTime[minuteKeyMinus] || [];
      const bucketPlus = topTradeQueueByDateTime[minuteKeyPlus] || [];

      debug.counts.minuteBucketMinus = bucketMinus.length;
      debug.counts.minuteBucketPlus = bucketPlus.length;

      function bestDeltaMsForBucket(bucket) {
        let best = 999999999;
        for (let i = 0; i < bucket.length; i++) {
          const it = bucket[i];
          if (!it) continue;
          if (String(it.topSym || "").toUpperCase() !== s) continue;
          if (Number(it.topAbsQty) !== q) continue;
          if (!(it.topTs instanceof Date) || isNaN(it.topTs.getTime()))
            continue;
          const d = Math.abs(it.topTs.getTime() - tradeTs.getTime());
          if (d < best) best = d;
        }
        return best;
      }

      const dMinus = bestDeltaMsForBucket(bucketMinus);
      const dPlus = bestDeltaMsForBucket(bucketPlus);

      const hasMinus = dMinus < 999999999;
      const hasPlus = dPlus < 999999999;

      if (hasMinus || hasPlus) {
        if (hasMinus && (!hasPlus || dMinus <= dPlus)) {
          minuteKeyUsed = minuteKeyMinus;
          dtBucket = bucketMinus;
          debug.why = "Base minute bucket empty; used -1 minute bucket.";
        } else {
          minuteKeyUsed = minuteKeyPlus;
          dtBucket = bucketPlus;
          debug.why = "Base minute bucket empty; used +1 minute bucket.";
        }
      } else {
        debug.why =
          "No TosTop TRD candidates in the same minute bucket (or +/- 1 minute buckets).";
        return { item: null, debug: debug };
      }
    } else {
      debug.why = "Using base minute bucket.";
    }

    const candidates = [];
    for (let i = 0; i < dtBucket.length; i++) {
      const it = dtBucket[i];
      if (!it) continue;
      if (String(it.topSym).toUpperCase() !== s) continue;
      if (Number(it.topAbsQty) !== q) continue;
      candidates.push({ idx: i, it: it });
    }

    debug.counts.symQtyCandidates = candidates.length;

    // Partial fills: TosTrades has the order total (e.g. -2, -76, -50).
    // TosTop often splits that into 2+ TRD rows (e.g. -1 and -1, or -10 and -66).
    // Those pieces can land in the same second, a few seconds apart, or across
    // the minute boundary. Do not treat that as "symbol missing in this minute."
    if (!candidates.length) {
      const partial = pullPartialFillEnrichment_(
        al,
        dateIso,
        tradeTs,
        minuteKey,
        s,
        q,
        matchPrice,
        debug,
      );
      if (partial && partial.item) {
        return partial;
      }

      debug.why =
        "Minute bucket had TRD rows, but none matched symbol+qty (and no partial-fill qty sum).";
      return { item: null, debug: debug };
    }

    const secondsUnknown =
      tradeTs instanceof Date &&
      tradeTs.getSeconds &&
      tradeTs.getSeconds() === 0;

    candidates.sort((a, b) => {
      if (secondsUnknown) {
        const pa = priceDelta(a.it);
        const pb = priceDelta(b.it);
        if (pa !== pb) return pa - pb;
        const ta = timeDeltaMs(a.it);
        const tb = timeDeltaMs(b.it);
        if (ta !== tb) return ta - tb;
        return a.idx - b.idx;
      }
      const ta = timeDeltaMs(a.it);
      const tb = timeDeltaMs(b.it);
      if (ta !== tb) return ta - tb;
      const pa = priceDelta(a.it);
      const pb = priceDelta(b.it);
      if (pa !== pb) return pa - pb;
      return a.idx - b.idx;
    });

    function priceDelta(it) {
      const a = toNum(it.topPrice);
      const b = toNum(matchPrice);
      if (isNaN(a) || isNaN(b)) return 999999;
      return Math.abs(a - b);
    }

    function timeDeltaMs(it) {
      if (!(it.topTs instanceof Date) || isNaN(it.topTs.getTime()))
        return 999999999;
      return Math.abs(it.topTs.getTime() - tradeTs.getTime());
    }

    if (expectedFillCount > 1 && candidates.length >= expectedFillCount) {
      let sumMiscFees = 0,
        sumFeesComm = 0,
        sumAmount = 0;
      let sawMiscFees = false,
        sawFeesComm = false,
        sawAmount = false;

      const picked = candidates.slice(0, expectedFillCount).map((x) => x.it);
      const idxs = candidates
        .slice(0, expectedFillCount)
        .map((x) => x.idx)
        .sort((a, b) => b - a);

      for (let k = 0; k < idxs.length; k++) {
        const removed = dtBucket.splice(idxs[k], 1)[0];
        removeFromExactIndex(removed);

        const mf = toNum(removed.miscFees);
        const fc = toNum(removed.feesComm);
        const am = toNum(removed.amount);

        if (!isNaN(mf)) {
          sumMiscFees += mf;
          sawMiscFees = true;
        }
        if (!isNaN(fc)) {
          sumFeesComm += fc;
          sawFeesComm = true;
        }
        if (!isNaN(am)) {
          sumAmount += am;
          sawAmount = true;
        }
      }

      const out = {
        miscFees: sawMiscFees ? sumMiscFees : "",
        feesComm: sawFeesComm ? sumFeesComm : "",
        amount: sawAmount ? sumAmount : "",
        matchedBy: "aggregateClosestInMinute",
        matchedKey: minuteKeyUsed,
      };
      debug.why = "Matched aggregate closest-in-time within minute bucket.";
      return { item: out, debug: debug };
    }

    const chosen = candidates[0];
    const picked = dtBucket.splice(chosen.idx, 1)[0];
    removeFromExactIndex(picked);

    picked.matchedBy = "closestInMinute";
    picked.matchedKey = minuteKeyUsed;
    debug.why = "Matched closest-in-time within minute bucket.";
    return { item: picked, debug: debug };
  }

  function pullTopTradeEnrichmentButterfly(
    Account,
    dateIso,
    timeHHmm,
    sym,
    descNeedle,
  ) {
    const needle = String(descNeedle || "BUTTERFLY")
      .trim()
      .toUpperCase();
    const acc = toStr(Account).trim().toUpperCase();
    const symU = String(sym || "")
      .trim()
      .toUpperCase();

    function consumeFromMinuteKey(dtKey, whyLabel) {
      const bucket = topTradeQueueByDateTime[dtKey] || [];
      const picks = [];

      for (let i = bucket.length - 1; i >= 0; i--) {
        const it = bucket[i];
        if (!it) continue;
        if (String(it.topSym).trim().toUpperCase() !== symU) continue;
        const descU = String(it.topDesc || "").toUpperCase();
        if (descU.indexOf(needle) < 0) continue;

        bucket.splice(i, 1);
        removeFromExactIndex(it);
        picks.push(it);
      }

      if (!picks.length) return null;

      function addMaybe(sum, v) {
        if (v === null || v === undefined || toStr(v).trim() === "") return sum;
        const n = toNum(v);
        if (isNaN(n)) return sum;
        return sum + n;
      }

      let sumMiscFees = 0,
        sumFeesComm = 0,
        sumAmount = 0;
      let sawAnyMiscFees = false,
        sawAnyFeesComm = false,
        sawAnyAmount = false;

      for (let i = 0; i < picks.length; i++) {
        const it = picks[i];
        if (toStr(it.miscFees).trim() !== "") sawAnyMiscFees = true;
        if (toStr(it.feesComm).trim() !== "") sawAnyFeesComm = true;
        if (toStr(it.amount).trim() !== "") sawAnyAmount = true;

        sumMiscFees = addMaybe(sumMiscFees, it.miscFees);
        sumFeesComm = addMaybe(sumFeesComm, it.feesComm);
        sumAmount = addMaybe(sumAmount, it.amount);
      }

      return {
        item: {
          miscFees: sawAnyMiscFees ? sumMiscFees : "",
          feesComm: sawAnyFeesComm ? sumFeesComm : "",
          amount: sawAnyAmount ? sumAmount : "",
          matchedBy: "butterflySumMinuteBucket",
          matchedKey: dtKey,
          pickedCount: picks.length,
        },
        debug: {
          why: whyLabel,
          dtKey: dtKey,
          sym: symU,
          pickedCount: picks.length,
        },
      };
    }

    const dtKey0 = [acc, dateIso, timeHHmm].join("|");
    const res0 = consumeFromMinuteKey(
      dtKey0,
      "Summed butterfly legs from exact minute bucket.",
    );
    if (res0) return res0;

    const baseTs = toDateObject(dateIso, timeHHmm);
    if (baseTs instanceof Date && !isNaN(baseTs.getTime())) {
      const minusTs = new Date(baseTs.getTime() - 60 * 1000);
      const plusTs = new Date(baseTs.getTime() + 60 * 1000);
      const minusKey = [
        acc,
        normalizeDate(minusTs),
        normalizeTime(minusTs),
      ].join("|");
      const plusKey = [acc, normalizeDate(plusTs), normalizeTime(plusTs)].join(
        "|",
      );

      if (
        (topTradeQueueByDateTime[minusKey] || []).some(
          (it) =>
            String(it.topSym || "")
              .trim()
              .toUpperCase() === symU,
        )
      ) {
        const resMinus = consumeFromMinuteKey(
          minusKey,
          "Summed butterfly legs from -1 minute bucket.",
        );
        if (resMinus) return resMinus;
      }
      if (
        (topTradeQueueByDateTime[plusKey] || []).some(
          (it) =>
            String(it.topSym || "")
              .trim()
              .toUpperCase() === symU,
        )
      ) {
        const resPlus = consumeFromMinuteKey(
          plusKey,
          "Summed butterfly legs from +1 minute bucket.",
        );
        if (resPlus) return resPlus;
      }
    }

    return {
      item: null,
      debug: {
        why: "No TosTop TRD candidates for butterfly symbol in minute (or adjacent minutes).",
        dtKey: dtKey0,
        sym: symU,
      },
    };
  }

  /**
   * IRON CONDOR: TosTop prints ONE TRD for the whole condor
   *   "SOLD -1 IRON CONDOR SPX … @1.20"
   * TosTrades prints four legs. Generic pull treats qty 1 + symbol as a
   * single-leg fill and removes that TRD, so later IC SELL groups in the
   * same minute log NO_TRD_ROWS.
   *
   * This helper runs FIRST. It pulls one unused TosTop row whose
   * DESCRIPTION contains "IRON CONDOR" and whose parsed underlying
   * matches. Two packages in the same minute (MSFT 09:43 x2) each get
   * their own TRD.
   */
  function pullTopTradeEnrichmentIronCondor(
    Account,
    dateIso,
    timeHHmm,
    sym,
    tradeTs,
    wantQty,
  ) {
    const acc = toStr(Account).trim().toUpperCase();
    const symU = String(sym || "")
      .trim()
      .toUpperCase();

    function pickFromMinuteKey(dtKey) {
      const bucket = topTradeQueueByDateTime[dtKey] || [];
      let bestIdx = -1;
      let bestDist = 999999999;

      for (let i = 0; i < bucket.length; i++) {
        const it = bucket[i];
        if (!it) continue;
        if (
          String(it.topSym || "")
            .trim()
            .toUpperCase() !== symU
        )
          continue;
        const descU = String(it.topDesc || "").toUpperCase();
        if (descU.indexOf("IRON CONDOR") < 0) continue;

        const want = Number(wantQty);
        const rowQty = Number(it.topAbsQty);
        if (isFinite(want) && want > 0 && isFinite(rowQty) && rowQty !== want)
          continue;

        let dist = 0;
        if (
          tradeTs instanceof Date &&
          !isNaN(tradeTs.getTime()) &&
          it.topTs instanceof Date &&
          !isNaN(it.topTs.getTime())
        ) {
          dist = Math.abs(it.topTs.getTime() - tradeTs.getTime());
        }
        if (bestIdx < 0 || dist < bestDist) {
          bestIdx = i;
          bestDist = dist;
        }
      }

      if (bestIdx < 0) return null;
      const picked = bucket.splice(bestIdx, 1)[0];
      removeFromExactIndex(picked);
      return picked;
    }

    const debug = {
      requested: {
        Account: acc,
        dateIso: dateIso,
        timeHHmm: timeHHmm,
        sym: symU,
      },
      why: "",
    };

    const dtKey0 = [acc, dateIso, timeHHmm].join("|");
    let picked = pickFromMinuteKey(dtKey0);
    let minuteKeyUsed = dtKey0;

    if (!picked) {
      const baseTs = toDateObject(dateIso, timeHHmm);
      if (baseTs instanceof Date && !isNaN(baseTs.getTime())) {
        const minusTs = new Date(baseTs.getTime() - 60 * 1000);
        const plusTs = new Date(baseTs.getTime() + 60 * 1000);
        const minusKey = [
          acc,
          normalizeDate(minusTs),
          normalizeTime(minusTs),
        ].join("|");
        const plusKey = [
          acc,
          normalizeDate(plusTs),
          normalizeTime(plusTs),
        ].join("|");

        picked = pickFromMinuteKey(minusKey);
        if (picked) minuteKeyUsed = minusKey;
        if (!picked) {
          picked = pickFromMinuteKey(plusKey);
          if (picked) minuteKeyUsed = plusKey;
        }
      }
    }

    if (!picked) {
      debug.why =
        "No TosTop IRON CONDOR TRD for this symbol in minute (or ±1).";
      debug.dtKey = dtKey0;
      return { item: null, debug: debug };
    }

    picked.matchedBy = "ironCondorPackageMinute";
    picked.matchedKey = minuteKeyUsed;
    debug.why = "Matched one TosTop IRON CONDOR package TRD.";
    debug.dtKey = minuteKeyUsed;
    return { item: picked, debug: debug };
  }

  return {
    previewTopCandidates: previewTopCandidates,
    removeFromExactIndex: removeFromExactIndex,
    pullPartialFillEnrichment_: pullPartialFillEnrichment_,
    pullTopTradeEnrichment: pullTopTradeEnrichment,
    pullTopTradeEnrichmentButterfly: pullTopTradeEnrichmentButterfly,
    pullTopTradeEnrichmentIronCondor: pullTopTradeEnrichmentIronCondor,
  };
}
