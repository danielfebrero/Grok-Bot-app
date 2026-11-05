---
name: site-playbooks-delta
description: >-
  Delta site playbook: fares or booking a flight; not check-in, changes, or
  awards.
---
# Site playbook: Delta

Read this before you dispatch a computerUse subagent to this site. Follow the Fastest path, fill in every URL you know, then paste the Dispatch snippet into the Task text as written. The subagent does not read skills, so every rule it needs is in the snippet, including the ground rules at its end. Secrets never go in the Task text. They come through `RequestUserForm`.

Jobs in this file: Delta flight search; Delta booking a picked flight.

## Delta flight search

### When to use

The user wants Delta's cash fares and schedules for a route and date. The parent needs origin and destination IATA codes, dates, the count of seated travelers, and cabin. The user gets each outbound flight's times, duration, stops, flight numbers, and the price in each fare column. Multi-carrier fare shopping belongs in the Google Flights playbook.

### Fastest path

`browser_navigate` the booking deep link. It lands on the booking form prefilled from the params, never on results, so one press of Find Flights is needed. Results then load on a session-bound `/flightsearch/search-results` URL.

```
https://www.delta.com/flightsearch/book-a-flight
  ?tripType={ONE_WAY|ROUND_TRIP}
  &originCity={ORIGIN_IATA}
  &destinationCity={DEST_IATA}
  &departureDate={DEPART_YYYY-MM-DD}
  &returnDate={RETURN_YYYY-MM-DD}      # omit for one-way
  &paxCount={SEATED_TRAVELERS}         # adults and children with a seat, 1 to 9
  &cabinFareClass={DCP|DPPS|FIRST|D1}  # omit for Main
```

`paxCount` counts every traveler with a seat, children and infants in a seat included; seniors count as adults. The form has only a passenger count, so everyone is searched as an adult; children's ages and lap infants go on the passenger page. `cabinFareClass` codes: `DCP` Comfort, `DPPS` Premium Select, `FIRST` First, `D1` Delta One. With one set, results show only that cabin's column.

### Dispatch snippet

```
Open https://www.delta.com/flightsearch/book-a-flight?tripType=<ONE_WAY|ROUND_TRIP>&originCity=<ORIGIN>&destinationCity=<DEST>&departureDate=<YYYY-MM-DD>&returnDate=<YYYY-MM-DD or omit>&paxCount=<SEATED TRAVELERS>&cabinFareClass=<DCP, DPPS, FIRST, or D1; omit for Main or all cabins>
Wait 4 s, then snapshot. Dismiss a cookie banner or a SkyMiles sign-in prompt if it offers to continue without signing in. Check the form has <ORIGIN>, <DEST>, the dates, the trip type, and <SEATED TRAVELERS> passengers; fix any field that differs once, then press Find Flights once.
Wait until flight rows with prices appear (up to 10 s), then snapshot.
For each outbound row read depart and arrive times (+1 = next day), duration, stops with connecting airports, flight numbers (DL nnnn), and the price in each column. A sold-out or missing cell is not available.
If the page says no flights match, report an empty result. If the grid stays empty after the wait, report empty_grid with the URL and stop.
Return a table sorted by price plus the page URL and one screenshot path. Read only. Never click a fare or Continue.
If you see Access Denied, a captcha, a sign-in that cannot be skipped, a password, 2FA, or payment wall, stop and report the URL.
Ground rules for this Task. A Timeout or a net::ERR_* code from browser_navigate means the box cannot reach that host. Do not retry it, even if the tool result says to, and do not run DNS or connectivity checks. Move to the next source this Task names, or report the code and URL. An HTTP block (403, 419, 429, 503) follows this Task's own rule. Work in the tab you start in. Do not open a second tab, and switch tabs only when a tool result names a page you did not intend. In-page tabs are part of the page. A page script is an async (page) => function: run it with browser_run_code if you hold it. Without it, read rendered text from browser_snapshot or browser_find, read a <script> tag from the HTML that curl fetches in the box shell, and report a field neither yields as unreadable, never visible text in its place. Name the read you used. Snapshot a page state once. After a click, a scroll, or a wait this Task calls for, take one new snapshot. Do not snapshot the same state twice. Report as soon as the stop rule above is met, one line per item, only the fields asked for, no page dumps.
```

### Read the result

Columns are Main, Comfort, First, and Delta One, with Premium Select on routes that have it; no Basic column shows by default. Prices are per traveler unless the page says total, and round trip cells read "From $X Round Trip". Round trip pages show outbound rows first; return rows render only after an outbound is picked, which this search must not do.

### When blocked

- Akamai fronts delta.com. In a browser that exposes automation the form loads, but Find Flights returns an "Access Denied" page with a Reference number. Other tells are a captcha or a results page that never fills.
- After a block or `empty_grid`, do not dispatch the deep link again. Hand the route to the Google Flights playbook and read Delta's rows there.
- If the page asks for a password, 2FA code, captcha, or payment, stop. Do not solve or bypass. Report which wall appeared and the URL it was on.

## Delta booking a picked flight

### When to use

The user picked a Delta flight, from FlightSearch, a card, or the search section, and none of your tools books it in chat. The parent needs the flight numbers and departure time of each leg, the dates, the fare column, and the count of seated travelers. The search takes only a seated-traveler count, so the parent gives children's ages and adds lap infants on the passenger page. Fill `<CHECKOUT LINK>` with the `checkoutUrl` in the user's Book request only when its host is delta.com or a subdomain of it; otherwise leave the slot as is. The subagent selects those fares and stops at the first page that asks for travelers, a sign-in, or payment. The parent fills that page as the `flight-booking` skill says and resumes the same subagent.

### Fastest path

Start from `<CHECKOUT LINK>` when it is filled in; it should open on those exact flights. When it is not filled in, or opens neither a trip summary nor the search page, start from the search section's deep link with the same route, dates, trip type, `paxCount`, and the chosen column's `cabinFareClass`, and press Find Flights once. Selecting the flight and fare then takes clicks.

### Dispatch snippet

```
If <CHECKOUT LINK> is a delta.com link, open it and dismiss any cookie banner. On a trip summary with flights <DL flight numbers> and the <fare column> fare already selected, go to the summary-page line below. On a delta.com search form or results, stay there and skip the Open line. On anything else, also report checkout_link_failed with its URL.
Open https://www.delta.com/flightsearch/book-a-flight?tripType=<ONE_WAY|ROUND_TRIP>&originCity=<ORIGIN>&destinationCity=<DEST>&departureDate=<YYYY-MM-DD>&returnDate=<YYYY-MM-DD or omit>&paxCount=<SEATED TRAVELERS>&cabinFareClass=<DCP, DPPS, FIRST, or D1; omit for Main or all cabins>
Wait 4 s, then snapshot. Dismiss a SkyMiles prompt that offers to continue without signing in. Check <ORIGIN>, <DEST>, the dates, and <SEATED TRAVELERS> passengers, fix what differs once, press Find Flights once, and wait for flight rows (up to 10 s). Leave lap infants for the passenger page.
If the outbound row for <DL flight numbers> at <HH:MM> is missing or its <fare column> cell is sold out, report fare_not_available with the URL and stop.
Click that price and confirm the selection. For a round trip, pick the return row for <DL flight numbers> the same way.
Continue past each summary page, declining seats, bags, upgrades, and insurance.
Stop at the first page that asks for traveler details, a sign-in, or payment. Report its URL, the fields it shows, the flights and fares selected, and the total shown, with one screenshot path.
If the total isn't for <SEATED TRAVELERS> seated travelers or a flight differs from the ones above, report the difference and stop. If you see Access Denied, a captcha, password, or 2FA prompt, stop and report the URL.
Ground rules for this Task. A Timeout or a net::ERR_* code from browser_navigate means the box cannot reach that host. Do not retry it, even if the tool result says to, and do not run DNS or connectivity checks. Move to the next source this Task names, or report the code and URL. An HTTP block (403, 419, 429, 503) follows this Task's own rule. Work in the tab you start in. Do not open a second tab, and switch tabs only when a tool result names a page you did not intend. In-page tabs are part of the page. A page script is an async (page) => function: run it with browser_run_code if you hold it. Without it, read rendered text from browser_snapshot or browser_find, read a <script> tag from the HTML that curl fetches in the box shell, and report a field neither yields as unreadable, never visible text in its place. Name the read you used. Snapshot a page state once. After a click, a scroll, or a wait this Task calls for, take one new snapshot. Do not snapshot the same state twice. Report as soon as the stop rule above is met, one line per item, only the fields asked for, no page dumps.
```

### When blocked

- Same Akamai tells as the search section.
- After a block, a site error, a stop code, or the user skipping a help request, do not dispatch again. Send the user the link the `flight-booking` skill's stopped-booking step names.
- If the page asks for a password, 2FA code, captcha, or payment, stop. Do not solve or bypass. Report which wall appeared and the URL it was on.

## Working with the subagent

Recipe steps name actions, not tools. A page script, a scroll, or a fill lands on whichever browser tools the subagent holds, and the ground rules say how, so paste the snippet as written. A subagent without a page-script tool reads what the snapshot or a shell fetch shows and names that read in its report.

Fill in every URL you know before you dispatch. A dispatch that changes something on the site (a cart, a booking, a form) never goes out without its URL. A read-only lookup may start from a name and a place when its section allows it. When the Fastest path is a shell call, run it with a short timeout and never queue a second lookup behind it. A failed call follows the section's own rule for that response (a single retry, a fresh key, or stop and report); when the section has no rule for it, or the call hangs, dispatch the subagent at once. When the section says the call is unreliable from the box, dispatch in the same round and stop the subagent with `StopSubagent` if the call returns first. A change to a job in flight goes to the same subagent, with `MessageSubagent` while it runs and `Task` with `resume` after it reports. A fresh dispatch spends rounds finding a tab and re-reading the page before it does anything.
