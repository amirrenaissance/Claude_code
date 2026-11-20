-- Are setter bookings net new, or meetings email (inbox managers) would have booked anyway?
-- Cohort: funding positive replies entering the GoFunnel positive-reply list Nov 10-14 2026.
-- Compares booking rates for leads setters reached live vs never reached; matches bookings by phone or email.
WITH m AS (
  SELECT lead_key, any(lead_id) lead_id, right(replaceRegexpAll(any(phone),'\\D',''),10) p10, lower(trim(any(email))) em, min(added_at) added
  FROM main.raw_gofunnel_dialer_member WHERE root_list_name = 'Positive Reply List November 4th'
  GROUP BY lead_key HAVING added >= '2026-11-10' AND added < '2026-11-15'),
b AS (
  SELECT created_at, channel, right(replaceRegexpAll(ifNull(phone,''),'\\D',''),10) p10, lower(trim(ifNull(email,''))) em
  FROM sb_portal.im_bookings WHERE offer='Funding' AND deleted_at IS NULL AND created_at >= '2026-11-09'),
j AS (SELECT m.lead_key, b.channel, b.created_at, m.added FROM m INNER JOIN b
      ON (m.p10 != '' AND m.p10 = b.p10) OR (m.em != '' AND m.em = b.em) WHERE b.created_at >= m.added - INTERVAL 1 DAY),
per AS (SELECT lead_key, max(channel='Call') call_b, max(channel='Email') email_b FROM j GROUP BY lead_key),
conn AS (SELECT DISTINCT lead_id FROM main.raw_gofunnel_call WHERE scheduled_at >= '2026-11-10' AND status NOT IN ('completed','no_answer','voicemail','gfmiss')),
t AS (SELECT m.lead_id IN (SELECT lead_id FROM conn) reached, ifNull(per.call_b,0) call_b, ifNull(per.email_b,0) email_b FROM m LEFT JOIN per ON m.lead_key = per.lead_key)
SELECT reached, count() leads, sum(call_b) call_booked, sum(email_b) email_booked FROM t GROUP BY reached;
-- Net-new share = (reached booked total - reached leads x unreached email rate) / reached call bookings
