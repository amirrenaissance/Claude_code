-- Stage 6 : meetings delivered per Partner key (Partners_Profile).
SELECT coalesce(nullIf(partner_key, ''), lower(partner)) AS pk, count() AS meetings,
       min(meeting_date) AS first, max(meeting_date) AS last, countIf(meeting_date >= today() - 30) AS m30
FROM core.meeting
WHERE ifNull(is_duplicate_of, '') = ''
GROUP BY pk
ORDER BY meetings DESC
