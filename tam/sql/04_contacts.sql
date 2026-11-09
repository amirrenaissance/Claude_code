-- Stage 4 -> amir_lab.tam_contacts : best {per_domain} contacts per domain, ranked by seniority,
-- then verified email, then phone on file. Values are passed through exactly as stored; nothing is guessed.
SELECT d, e, fn, ln, substring(ifNull(job_title, ''), 1, 50) AS t, sr, vs, ph,
       replaceRegexpOne(li, '^https?://(www\\.)?linkedin\\.com/in/', '') AS li
FROM (
    SELECT lower(splitByChar('@', email)[2]) AS d, lower(email) AS e, first_name AS fn, last_name AS ln, job_title,
           seniority AS sr, verification_status AS vs,
           if(ifNull(phone, '') != '', phone, ifNull(enriched_phone, '')) AS ph, ifNull(person_linkedin, '') AS li,
           row_number() OVER (PARTITION BY d ORDER BY
               multiIf(match(lower(concat(ifNull(seniority, ''), ' ', ifNull(job_title, ''))), 'owner|founder|c_suite|c-suite|c-team|\\bceo\\b|president|chief|managing (member|director|partner)'), 1,
                       match(lower(concat(ifNull(seniority, ''), ' ', ifNull(job_title, ''))), '\\bvp\\b|vice president|director|head'), 2, 3),
               verification_status IN ('valid', 'catch_all_valid') DESC, ph != '' DESC, e) AS rn
    FROM mirror.leads_current
    WHERE email LIKE '%@%' AND lower(splitByChar('@', email)[2]) IN (SELECT domain FROM {universe})
)
WHERE rn <= {per_domain}
