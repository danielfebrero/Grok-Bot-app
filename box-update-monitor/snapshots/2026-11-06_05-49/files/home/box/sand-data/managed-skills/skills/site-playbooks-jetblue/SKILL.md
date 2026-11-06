---
name: site-playbooks-jetblue
description: >-
  JetBlue site playbook: fares or booking a flight; not check-in, changes, or
  awards.
---
# Site playbook: JetBlue

Read this before you dispatch a computerUse subagent to this site. Follow the Fastest path, fill in every URL you know, then paste the Dispatch snippet into the Task text as written. The subagent does not read skills, so every rule it needs is in the snippet, including the ground rules at its end. Secrets never go in the Task text. They come through `RequestUserForm`.

Jobs in this file: JetBlue flight search; JetBlue booking a picked flight.

## JetBlue flight search

### When to use

The user wants JetBlue's cash fares and schedules for a route and date. The parent needs origin and destination IATA codes, dates, adult, child, and lap infant counts, and cabin. The user gets each outbound flight's times, duration, stops, flight numbers, and the price in each fare column, Mint included. Multi-carrier fare shopping belongs in the Google Flights playbook.

### Fastest path

`browser_navigate` the results deep link. It skips the home page form, and the page keeps the deep link as its URL, so it can be shared.

```
https://www.jetblue.com/booking/flights
  ?from={ORIGIN_IATA}
  &to={DEST_IATA}
  &depart={DEPART_YYYY-MM-DD}
  &return={RETURN_YYYY-MM-DD}          # omit for one-way
  &isMultiCity=false
  &noOfRoute=1
  &adults={ADULTS}                     # seniors count as adults
  &children={CHILDREN}
  &infants={LAP_INFANTS}
  &sharedMarket=false
  &roundTripFaresFlag=false            # prices stay one-way per person either way
  &usePoints=false                     # true prices in TrueBlue points
```

`infants` counts lap infants. If the page shows other dates or counts, fix them in the search form once. There is no cabin parameter; Mint is a fare column.

### Dispatch snippet

```
Open https://www.jetblue.com/booking/flights?from=<ORIGIN>&to=<DEST>&depart=<YYYY-MM-DD>&return=<RETURN or omit>&isMultiCity=false&noOfRoute=1&adults=<ADULTS>&children=<CHILDREN>&infants=<LAP INFANTS>&sharedMarket=false&roundTripFaresFlag=false&usePoints=false
Wait 4 s, then snapshot. Check the page shows <ORIGIN>, <DEST>, the dates, and the passenger counts. If it shows a form or a field differs, fix it once and search once.
Wait until flight rows with prices appear (up to 10 s), then snapshot. Read the fare column names off the page header, usually Main, EvenMore, and Mint, and include every column shown.
For each outbound row read depart and arrive times (+1 = next day), duration, stops with connecting airports, flight numbers (B6 nnnn), and the price in each column. A sold-out or missing cell is not available.
If the page says no flights match, report an empty result. If the grid stays empty after the wait, report empty_grid with the URL and stop.
Return a table sorted by price plus the page URL and one screenshot path. Read only. Never click a fare or Continue.
If you see Access Denied, a Client Challenge, a captcha, sign-in or password, 2FA, or payment wall, stop and report the URL.
Ground rules for this Task. A Timeout or a net::ERR_* code from browser_navigate means the box cannot reach that host. Do not retry it, even if the tool result says to, and do not run DNS or connectivity checks. Move to the next source this Task names, or report the code and URL. An HTTP block (403, 419, 429, 503) follows this Task's own rule. Work in the tab you start in. Do not open a second tab, and switch tabs only when a tool result names a page you did not intend. In-page tabs are part of the page. A page script is an async (page) => function: run it with browser_run_code if you hold it. Without it, read rendered text from browser_snapshot or browser_find, read a <script> tag from the HTML that curl fetches in the box shell, and report a field neither yields as unreadable, never visible text in its place. Name the read you used. Snapshot a page state once. After a click, a scroll, or a wait this Task calls for, take one new snapshot. Do not snapshot the same state twice. Report as soon as the stop rule above is met, one line per item, only the fields asked for, no page dumps.
```

### Read the result

On the routes checked, the columns were Main (standard seat), EvenMore (extra legroom), and Mint (lie-flat, on some routes); take the names from the header. Each cell shows the cheapest fare in its column. Clicking a cell opens tier cards, which this search does not open; the Main column opened Main Base, Main, and Main Flex. Prices are one-way per person, round trips included. Round trip pages show outbound rows first; return rows render only after an outbound is picked, which this search must not do.

### When blocked

- jetblue.com runs a client challenge. A browser that exposes automation gets a "Client Challenge" page that says "Enter the characters seen in the image". That was a live check from a residential connection, not yet from a box. Other tells are Access Denied or a results page that never fills.
- After a block or `empty_grid`, do not dispatch the deep link again. Hand the route to the Google Flights playbook and read JetBlue's rows there.
- If the page asks for a password, 2FA code, captcha, or payment, stop. Do not solve or bypass. Report which wall appeared and the URL it was on.

## JetBlue booking a picked flight

### When to use

The user picked a JetBlue flight, from FlightSearch, a card, or the search section, and none of your tools books it in chat. The parent needs the flight numbers and departure time of each leg, the dates, the fare column, the adult, child, and lap infant counts, and each child's and lap infant's age. The search link carries the child and lap infant counts, and the parent gives their ages on the traveler page. Fill `<CHECKOUT LINK>` with the `checkoutUrl` in the user's Book request only when its host is jetblue.com or a subdomain of it; otherwise leave the slot as is. The subagent selects those fares and stops at the first page that asks for travelers, a sign-in, or payment. The parent fills that page as the `flight-booking` skill says and resumes the same subagent.

The search reads only the cheapest fare in each column, so the subagent selects the cheapest tier card in the picked column and reports its name; in the Main column that was Main Base. The Main column also holds a tier named Main, so name the tier the subagent reports, not the column, to the user before the payment widget.

### Fastest path

Start from `<CHECKOUT LINK>` when it is filled in; it should open on those exact flights. When it is not filled in, or opens neither a trip summary nor the search page, start from the search section's deep link with the same route, dates, and passenger counts. There is no deep link past the results, so selecting the flight and fare takes clicks.

### Dispatch snippet

```
If <CHECKOUT LINK> is a jetblue.com link, open it and dismiss any cookie banner. On a trip summary with flights <B6 flight numbers> and a fare in the <fare column> column already selected, go to the summary-page line below. On a jetblue.com search form or results, stay there and skip the Open line. On anything else, also report checkout_link_failed with its URL.
Open https://www.jetblue.com/booking/flights?from=<ORIGIN>&to=<DEST>&depart=<YYYY-MM-DD>&return=<RETURN or omit>&isMultiCity=false&noOfRoute=1&adults=<ADULTS>&children=<CHILDREN>&infants=<LAP INFANTS>&sharedMarket=false&roundTripFaresFlag=false&usePoints=false
Wait 4 s, then snapshot. Check the route, dates, and the <ADULTS> adult, <CHILDREN> child, and <LAP INFANTS> lap infant counts. If one differs, fix it in the search form once and search once; if it won't take the children or lap infants, report children_not_set with the URL and stop. Wait for flight rows (up to 10 s).
If the outbound row for <B6 flight numbers> at <HH:MM> is missing or its <fare column> cell is sold out, report fare_not_available with the URL and stop.
Click that cell; if tier cards open, select the cheapest and report its name. Pick a return row the same way.
Continue past each summary page, declining seats, bags, upgrades, and insurance.
Stop at the first page that asks for traveler details, a sign-in, or payment. Report its URL, the fields it shows, the flights and fares selected, and the total shown, with one screenshot path.
If the total isn't for <ADULTS + CHILDREN> seated travelers or a flight differs from the ones above, report the difference and stop. If you see Access Denied, a Client Challenge, a captcha, password, or 2FA prompt, stop and report the URL.
Ground rules for this Task. A Timeout or a net::ERR_* code from browser_navigate means the box cannot reach that host. Do not retry it, even if the tool result says to, and do not run DNS or connectivity checks. Move to the next source this Task names, or report the code and URL. An HTTP block (403, 419, 429, 503) follows this Task's own rule. Work in the tab you start in. Do not open a second tab, and switch tabs only when a tool result names a page you did not intend. In-page tabs are part of the page. A page script is an async (page) => function: run it with browser_run_code if you hold it. Without it, read rendered text from browser_snapshot or browser_find, read a <script> tag from the HTML that curl fetches in the box shell, and report a field neither yields as unreadable, never visible text in its place. Name the read you used. Snapshot a page state once. After a click, a scroll, or a wait this Task calls for, take one new snapshot. Do not snapshot the same state twice. Report as soon as the stop rule above is met, one line per item, only the fields asked for, no page dumps.
```

### When blocked

- Same tells as the search section.
- After a block, a site error, a stop code, or the user skipping a help request, do not dispatch again. Send the user the link the `flight-booking` skill's stopped-booking step names.
- If the page asks for a password, 2FA code, captcha, or payment, stop. Do not solve or bypass. Report which wall appeared and the URL it was on.

## Working with the subagent

Recipe steps name actions, not tools. A page script, a scroll, or a fill lands on whichever browser tools the subagent holds, and the ground rules say how, so paste the snippet as written. A subagent without a page-script tool reads what the snapshot or a shell fetch shows and names that read in its report.

Fill in every URL you know before you dispatch. A dispatch that changes something on the site (a cart, a booking, a form) never goes out without its URL. A read-only lookup may start from a name and a place when its section allows it. When the Fastest path is a shell call, run it with a short timeout and never queue a second lookup behind it. A failed call follows the section's own rule for that response (a single retry, a fresh key, or stop and report); when the section has no rule for it, or the call hangs, dispatch the subagent at once. When the section says the call is unreliable from the box, dispatch in the same round and stop the subagent with `StopSubagent` if the call returns first. A change to a job in flight goes to the same subagent, with `MessageSubagent` while it runs and `Task` with `resume` after it reports. A fresh dispatch spends rounds finding a tab and re-reading the page before it does anything.
