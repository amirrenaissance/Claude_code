-- Positive-reply supply ladder: dials, agent-handled conversations and bookings by tier.
-- Tier = time between the lead entering the positive-reply list and the call (P0 <1h, P1 1-24h,
-- P2 1-3d, P3 3d+); carry-over lists are P4. Funding calls, Nov 10-18 2026.
-- Used for the 'Supply ladder' table on the Inputs tab of Renaissance_Setter_PnL.xlsx.
WITH m AS (
  SELECT lead_key, any(lead_id) lead_id, argMin(root_list_name, added_at) first_list, min(added_at) added
  FROM main.raw_gofunnel_dialer_member
  WHERE root_list_name IN ('Positive Reply List November 4th','Funding carry-over recent','Funding carry-over old')
  GROUP BY lead_key),
k AS (SELECT c.lead_id, c.scheduled_at, c.status, m.first_list, m.added
      FROM main.raw_gofunnel_call c JOIN m ON c.lead_id = m.lead_id
      WHERE c.scheduled_at >= '2026-11-10' AND c.scheduled_at < '2026-11-19'),
t AS (SELECT *, multiIf(first_list='Funding carry-over old','P4 old carry-over', first_list='Funding carry-over recent','P4 recent carry-over',
        added < '2026-11-10','P3 (3d+)', dateDiff('minute', added, scheduled_at) <= 60,'P0 (<1h)',
        dateDiff('minute', added, scheduled_at) <= 1440,'P1 (1-24h)', dateDiff('minute', added, scheduled_at) <= 4320,'P2 (1-3d)','P3 (3d+)') tier
      FROM k WHERE scheduled_at >= added OR first_list != 'Positive Reply List November 4th')
SELECT tier, count() dials,
       countIf(status NOT IN ('completed','no_answer','voicemail','gfmiss')) handled,
       countIf(status = 'gfbook') booked
FROM t GROUP BY tier ORDER BY tier;
