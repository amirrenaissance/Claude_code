-- Stage 5 -> amir_lab.tam_engagement : human replies (and strict-positive replies) from people at each domain.
-- Note: these are replies to any of our campaigns, mostly borrower-side "Business Funding" campaigns.
SELECT lower(splitByChar('@', r.lead_email)[2]) AS d, count() AS n, countIf(p.is_positive = 1) AS pos,
       toDate(max(r.reply_timestamp)) AS last
FROM derived.v_reply_canonical r
LEFT JOIN derived.reply_is_positive_strict p ON toString(p.reply_id) = toString(r.reply_id)
WHERE NOT r.is_auto_reply AND lower(splitByChar('@', r.lead_email)[2]) IN (SELECT domain FROM {universe})
GROUP BY d
