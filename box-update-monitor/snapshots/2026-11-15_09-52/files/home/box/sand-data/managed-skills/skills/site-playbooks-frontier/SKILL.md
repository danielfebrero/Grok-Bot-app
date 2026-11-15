---
name: site-playbooks-frontier
description: 'Frontier site playbook: fares or booking a flight; not check-in or changes.'
---
# Site playbook: Frontier

Read this before you dispatch a computerUse subagent to this site. Follow the Fastest path, fill in every URL you know, then paste the Dispatch snippet into the Task text as written. The subagent does not read skills, so every rule it needs is in the snippet, including the ground rules at its end. Secrets never go in the Task text. They come through `RequestUserForm`.

Jobs in this file: Frontier flight search; Frontier booking a picked flight.

## Frontier flight search

### When to use

The user wants Frontier's fares and schedules for a route and date. The parent needs origin and destination IATA codes, dates, the adult count, and each child's and lap infant's age. The user gets each outbound flight's times, duration, stops, and the Standard price of each bundle, plus the Discount Den price for members. Multi-carrier fare shopping belongs in the Google Flights playbook.

### Fastest path

The select deep link redirects to a session-bound `booking.flyfrontier.com/Flight/Select` with the search applied. If it lands on an error or other dates or counts, fill the search form on `https://www.flyfrontier.com/` once.

```
https://booking.flyfrontier.com/Flight/InternalSelect
  ?o1={ORIGIN_IATA}
  &d1={DEST_IATA}
  &dd1={DEPART, YYYY-MM-DD}
  &dd2={RETURN, YYYY-MM-DD}&r=true     # omit both for one-way
  &ADT={ADULTS}                        # seniors count as adults
  &kids={CHILD_AGE}                    # once per child, ages 0-14; omit with no children
  &inl={LAP_INFANTS}                   # omit with no lap infants
  &ia={INFANT_AGE}                     # once per lap infant, 0 or 1
  &mon=true
```

Without `r=true` the return date is dropped. `CHD` and `INF` are ignored, so children go only in `kids` and lap infants only in `inl` and `ia`. Frontier has one cabin; extra legroom and UpFront Plus seats are bundle columns on each row.

### Dispatch snippet

```
Open https://booking.flyfrontier.com/Flight/InternalSelect?o1=<ORIGIN>&d1=<DEST>&dd1=<YYYY-MM-DD>&dd2=<RETURN>&r=true&ADT=<ADULTS>&kids=<AGE, once per child>&inl=<LAP INFANTS>&ia=<AGE, once per lap infant>&mon=true (omit dd2 and r one-way, kids with no children, inl and ia with no lap infants)
Wait for priced flight rows (up to 15 s). If none show or they are for another route, date, or count, fill the form at https://www.flyfrontier.com/ once (close any card popup first; give each child and lap infant an age), search, and wait again. If it won't take the children or lap infants, report children_not_set with the URL and stop.
For each outbound row read depart and arrive times (+1 = next day), duration, stops with connecting airports, and the Standard and Discount Den prices under Basic Fare, Economy Bundle, Premium Bundle, and Business Bundle. Flight numbers are only in a row's +Details; open it for up to the 5 cheapest rows, read F9 nnnn, and close it. Leave the rest blank.
If the page says no flights match, report an empty result. If the grid stays empty after the wait, report empty_grid with the URL and stop.
Return a table sorted by Basic Fare price plus the page URL and one screenshot path. Read only. Never click a price or Continue.
If you see Access Denied, a captcha or press-and-hold check, sign-in or password, 2FA, or payment wall, stop and report the URL.
Ground rules for this Task. A Timeout or a net::ERR_* code from browser_navigate means the box cannot reach that host. Do not retry it, even if the tool result says to, and do not run DNS or connectivity checks. Move to the next source this Task names, or report the code and URL. An HTTP block (403, 419, 429, 503) follows this Task's own rule. Work in the tab you start in. Do not open a second tab, and switch tabs only when a tool result names a page you did not intend. In-page tabs are part of the page. A page script is an async (page) => function: run it with browser_run_code if you hold it. Without it, read rendered text from browser_snapshot or browser_find, read a <script> tag from the HTML that curl fetches in the box shell, and report a field neither yields as unreadable, never visible text in its place. Name the read you used. Snapshot a page state once. After a click, a scroll, or a wait this Task calls for, take one new snapshot. Do not snapshot the same state twice. Report as soon as the stop rule above is met, one line per item, only the fields asked for, no page dumps.
```

### Read the result

Premium Bundle adds extra legroom and Business Bundle is UpFront Plus; for what each bundle includes, read the bundle descriptions on the page. The Discount Den price is for paid Discount Den members only. Quote the Standard price unless the user says they are a member, and pass the same choice to booking, which clicks the price you quoted. Flight numbers are left blank past the 5 cheapest rows; booking reads them from the row the user picks. Prices are per person, one-way or round-trip as the banner above the grid says.

### When blocked

- Tells. An Access Denied page, a captcha or press-and-hold check, or a spinner that never ends. DataDome loads on flyfrontier.com but did not block a live check from a residential connection, not yet from a box.
- After a block, `empty_grid`, or `children_not_set`, do not dispatch again. Hand the route to the Google Flights playbook and read Frontier's rows there.
- If the page asks for a password, 2FA code, captcha, or payment, stop. Do not solve or bypass. Report which wall appeared and the URL it was on.

## Frontier booking a picked flight

### When to use

The user picked a Frontier flight, from FlightSearch, a card, or the search section, and none of your tools books it in chat. The parent needs the flight numbers and departure time of each leg, the dates, the bundle column the user chose (Basic Fare unless they said otherwise), Standard or Discount Den (Discount Den only for a member, as the search quoted), the adult count, and each child's and lap infant's age. The search link carries each child's and lap infant's age. Fill `<CHECKOUT LINK>` with the `checkoutUrl` in the user's Book request only when its host is flyfrontier.com or a subdomain of it; otherwise leave the slot as is. Fill `<EXTRAS>` as the `flight-booking` skill's extras step says. The subagent selects those fares and stops at the first page that asks for travelers, a sign-in, or payment. The parent fills that page as the `flight-booking` skill says and resumes the same subagent.

Flight numbers are optional here; the subagent reads them in the picked row's +Details.

### Fastest path

Start from `<CHECKOUT LINK>` when it is filled in; it should open on those exact flights. When it is not filled in, or opens neither a trip summary nor the search page, start from the search section's deep link with the same route, dates, and travelers. Selecting the flight and bundle takes clicks.

### Dispatch snippet

```
If <CHECKOUT LINK> is a flyfrontier.com link, open it and dismiss any cookie banner. On a trip summary with <bundle column> fares already selected, go to the bundle-offer line below. On a flyfrontier.com search form or results, stay there and skip the Open line. On anything else, also report checkout_link_failed with its URL.
Open https://booking.flyfrontier.com/Flight/InternalSelect?o1=<ORIGIN>&d1=<DEST>&dd1=<YYYY-MM-DD>&dd2=<RETURN>&r=true&ADT=<ADULTS>&kids=<AGE, once per child>&inl=<LAP INFANTS>&ia=<AGE, once per lap infant>&mon=true (omit dd2 and r one-way, kids with no children, inl and ia with no lap infants)
Wait for priced flight rows (up to 15 s). If none show or they are for another route, date, or count, fill the form at https://www.flyfrontier.com/ once (close any card popup first; give each child and lap infant an age), search, and wait again. If it won't take the children or lap infants, report children_not_set with the URL and stop.
Read the F9 numbers in the +Details of the outbound row at <HH:MM>. If none departs then or they differ from <F9 flight numbers, if given>, report flight_not_found with the URL and stop. Click its <Standard or Discount Den> price under <bundle column>; pick a return row the same way.
Keep <bundle column> at any bundle offer.
Extras, before or after the traveler page: add exactly <EXTRAS> and decline every other offer, upgrade, and insurance. If one won't add, report extra_unavailable with the page's reason and continue.
Stop at the first page that asks for traveler details, a sign-in, or payment. Report its URL, the fields it shows, the flights and fares selected, each extra's price, and the total shown, with one screenshot path.
If the total isn't for <ADULTS + CHILDREN> seated travelers or a flight differs from the ones above, report the difference and stop. If you see Access Denied, a captcha or press-and-hold check, password, or 2FA prompt, stop and report the URL.
Ground rules for this Task. A Timeout or a net::ERR_* code from browser_navigate means the box cannot reach that host. Do not retry it, even if the tool result says to, and do not run DNS or connectivity checks. Move to the next source this Task names, or report the code and URL. An HTTP block (403, 419, 429, 503) follows this Task's own rule. Work in the tab you start in. Do not open a second tab, and switch tabs only when a tool result names a page you did not intend. In-page tabs are part of the page. A page script is an async (page) => function: run it with browser_run_code if you hold it. Without it, read rendered text from browser_snapshot or browser_find, read a <script> tag from the HTML that curl fetches in the box shell, and report a field neither yields as unreadable, never visible text in its place. Name the read you used. Snapshot a page state once. After a click, a scroll, or a wait this Task calls for, take one new snapshot. Do not snapshot the same state twice. Report as soon as the stop rule above is met, one line per item, only the fields asked for, no page dumps.
```

### When blocked

- Same tells as the search section.
- After a block, a site error, a stop code, or the user skipping a help request, do not dispatch again. Send the user the link the `flight-booking` skill's stopped-booking step names.
- If the page asks for a password, 2FA code, captcha, or payment, stop. Do not solve or bypass. Report which wall appeared and the URL it was on.

## Working with the subagent

Recipe steps name actions, not tools. A page script, a scroll, or a fill lands on whichever browser tools the subagent holds, and the ground rules say how, so paste the snippet as written. A subagent without a page-script tool reads what the snapshot or a shell fetch shows and names that read in its report.

Fill in every URL you know before you dispatch. A dispatch that changes something on the site (a cart, a booking, a form) never goes out without its URL. A read-only lookup may start from a name and a place when its section allows it. When the Fastest path is a shell call, run it with a short timeout and never queue a second lookup behind it. A failed call follows the section's own rule for that response (a single retry, a fresh key, or stop and report); when the section has no rule for it, or the call hangs, dispatch the subagent at once. When the section says the call is unreliable from the box, dispatch in the same round and stop the subagent with `StopSubagent` if the call returns first. A change to a job in flight goes to the same subagent, with `MessageSubagent` while it runs and `Task` with `resume` after it reports. A fresh dispatch spends rounds finding a tab and re-reading the page before it does anything.
