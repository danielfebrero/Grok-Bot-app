---
name: site-playbooks-air-canada
description: >-
  Air Canada site playbook: fares or booking a flight; not check-in, changes, or
  awards.
---
# Site playbook: Air Canada

Read this before you dispatch a computerUse subagent to this site. Follow the Fastest path, fill in every URL you know, then paste the Dispatch snippet into the Task text as written. The subagent does not read skills, so every rule it needs is in the snippet, including the ground rules at its end. Secrets never go in the Task text. They come through `RequestUserForm`.

Jobs in this file: Air Canada flight search; Air Canada booking a picked flight.

## Air Canada flight search

### When to use

The user wants Air Canada's cash fares and schedules for a route and date. The parent needs origin and destination IATA codes, dates, and the count of each traveler type: adults 16 and over, youth 12 to 15, children 2 to 11, lap infants, and infants in a seat. The user gets each outbound flight's times, duration, stops, and the price in each cabin, with flight numbers for the cheapest rows. Multi-carrier fare shopping belongs in the Google Flights playbook.

### Fastest path

`browser_navigate` the results link. It skips the home page form, opens on "Loading flight results", and moves to `/availability/ow/outbound`, or `/availability/rt/outbound` for a round trip, once the rows load.

```
https://www.aircanada.com/booking/us/en/aco/search
  ?org0={ORIGIN_IATA}
  &dest0={DEST_IATA}
  &orgType0=A
  &destType0=A
  &departureDate0={DEPART_DD/MM/YYYY}
  &org1={DEST_IATA}                    # round trip only
  &dest1={ORIGIN_IATA}                 # round trip only
  &orgType1=A                          # round trip only
  &destType1=A                         # round trip only
  &departureDate1={RETURN_DD/MM/YYYY}  # round trip only
  &adt={ADULTS}                        # 16 and over, seniors too
  &yth={YOUTH}                         # 12 to 15
  &chd={CHILDREN}                      # 2 to 11
  &inf={LAP_INFANTS}
  &ins={INFANTS_IN_SEAT}
  &marketCode={DOM|TBO|INT}
  &tripType={OneWay|RoundTrip}
  &isFlexible=false
```

Fill `<RETURN LEG>` with the five round trip parameters, `&org1=<DEST>&dest1=<ORIGIN>&orgType1=A&destType1=A&departureDate1=<RETURN DD/MM/YYYY>`, or leave it out for a one-way trip.

`marketCode` is `DOM` within Canada, `TBO` between Canada and the US, and `INT` for every other route. Send every other parameter, zeros included: a link that drops them opens an error page that says "Something went wrong" and that the session expired. The `us/en` path prices in US dollars. The Flight search summary counts every traveler, lap infants included.

### Dispatch snippet

```
Open https://www.aircanada.com/booking/us/en/aco/search?org0=<ORIGIN>&dest0=<DEST>&orgType0=A&destType0=A&departureDate0=<DD/MM/YYYY><RETURN LEG or omit>&adt=<ADULTS>&yth=<YOUTH>&chd=<CHILDREN>&inf=<LAP INFANTS>&ins=<INFANTS IN SEAT>&marketCode=<DOM|TBO|INT>&tripType=<OneWay|RoundTrip>&isFlexible=false
Wait until "Loading flight results" clears and flight rows appear (up to 30 s), then snapshot. Dismiss any cookie banner. If the page says "Something went wrong", report link_error with the URL and stop.
Check the Flight search summary shows <ORIGIN>, <DEST>, the dates, and <ALL TRAVELERS> passengers. If a field differs, fix it once from https://www.aircanada.com/home/us/en/aco/flights: in Departing from or Arriving in, clear the prefilled city, type the city, and pick the option that starts with the airport code; set Trip type and Passenger(s) in their menus; type dates as DD/MM and press Escape to close the calendar; then press Search in the Flights form.
For each outbound row read depart and arrive times (+1 = next day), duration, stops with connecting airports, and the price shown for each cabin. A cabin without a price is not available.
For the three cheapest rows, open Details for the flight numbers (AC nnnn) and the operating airline, then close it. When this task names a fare, such as Standard, also open those rows' Economy price, which only lists its fares, and read that fare's price.
If the page says no flights match, report an empty result. If no rows appear after the wait, report empty_grid with the URL and stop.
Return a table sorted by price plus the page URL and one screenshot path. Read only. Never select a fare or press Continue.
If you see Access Denied, a captcha, sign-in or password, 2FA, or payment wall, stop and report the URL.
Ground rules for this Task. A Timeout or a net::ERR_* code from browser_navigate means the box cannot reach that host. Do not retry it, even if the tool result says to, and do not run DNS or connectivity checks. Move to the next source this Task names, or report the code and URL. An HTTP block (403, 419, 429, 503) follows this Task's own rule. Work in the tab you start in. Do not open a second tab, and switch tabs only when a tool result names a page you did not intend. In-page tabs are part of the page. A page script is an async (page) => function: run it with browser_run_code if you hold it. Without it, read rendered text from browser_snapshot or browser_find, read a <script> tag from the HTML that curl fetches in the box shell, and report a field neither yields as unreadable, never visible text in its place. Name the read you used. Snapshot a page state once. After a click, a scroll, or a wait this Task calls for, take one new snapshot. Do not snapshot the same state twice. Report as soon as the stop rule above is met, one line per item, only the fields asked for, no page dumps.
```

### Read the result

Each row shows depart and arrive times, stops, duration, and a price for Economy, Premium Economy, and Business Class. A cabin's price is its cheapest fare. Opening the Economy price lists its fares: Basic on some routes, then Standard, Flex, Comfort, and Latitude. Prices are per person and include taxes; round trip prices are each way, based on the round-trip ticket. Round trip pages show outbound rows first; return rows render only after an outbound is picked, which this search must not do.

### When blocked

- Tells of a wall are Access Denied, a captcha, or a results page that never fills.
- After `link_error`, check the link against the Fastest path, add what is missing, and dispatch once more. After a block or `empty_grid`, do not dispatch the link again. Hand the route to the Google Flights playbook and read Air Canada's rows there.
- If the page asks for a password, 2FA code, captcha, or payment, stop. Do not solve or bypass. Report which wall appeared and the URL it was on.

## Air Canada booking a picked flight

### When to use

The user picked an Air Canada flight, from FlightSearch, a card, or the search section, and none of your tools books it in chat. The parent needs the flight numbers and departure time of each leg, the dates, the cabin and fare, such as Economy and Standard, and the count of each traveler type. The search link carries every traveler type as a count, and the passenger page asks each traveler's date of birth. Fill `<CHECKOUT LINK>` with the `checkoutUrl` in the user's Book request only when its host is aircanada.com or a subdomain of it; otherwise leave the slot as is. Fill `<EXTRAS>` as the `flight-booking` skill's extras step says. The subagent selects those fares and stops at the first page that asks for travelers, a sign-in, or payment. The parent fills that page as the `flight-booking` skill says and resumes the same subagent.

The trip review page shows a 10-minute "Time to book!" countdown, but the session ends 10 minutes after the current page opened, and each new page restarts that clock, so the whole booking can take longer. Typing on a page does not restart it. After 8 minutes on a page, an "Extend your session" dialog counts down the last 2, and its Continue restarts the clock. Nothing keeps the session alive while the subagent waits for a form, so a user who takes longer than 10 minutes on the passenger form or the payment form loses the cart. The resumed subagent then reopens the results, selects the flight and fare again, and enters the travelers it was given again; ask the user only for a field it reports missing. Anything typed on Payment is lost with the cart, so the payment form may need to go out again. The subagent waits on you at the passenger page and at Payment, so send each form as soon as it reports, and have every traveler's details ready before you dispatch when you can.

The passenger page asks each traveler, children included, for first, middle, and last name, gender (Female, Male, Undisclosed, or Unspecified), and date of birth (DD/MM/YYYY), an optional loyalty number and Known Traveler Number, and one email and phone number for the booking. It had no passport field on the routes checked.

After the passenger page come Seat selection (`/seatmap/0/0` on a nonstop one-way), Travel options (`/travel-options`), and Payment (`/payment`). None of them adds anything by default, and seats cost extra on the Standard fare. With a child aboard, Seat selection first opens a "Family seating" dialog saying children under 14 get seats near their adults at no charge when none are picked. "Skip seats" opens "Travel with ease with secured seats", whose Continue is a small link beside a large Select seats button. Travel options sells no bags, since extra bags are bought at check-in. No insurance offer appeared on the routes checked.

Payment asks for the card number, Expiry MM/YY, CVV, the cardholder's first and last name, and a billing address: Address Line 1, optional Address Line 2, Country/Region (preset to United States), City, Province/State, and Postal/ZIP code. It asks no email or phone. Credit Card / Debit Card / AC Credit is selected by default, and AC Wallet needs a sign-in. The Payment summary shows the Grand total in US dollars, and View price breakdown lists each tax. There is no consent checkbox: "I accept, purchase" accepts the fare rules and pays, so it is the final click.

### Fastest path

Start from `<CHECKOUT LINK>` when it is filled in; it should open on those exact flights. When it is not filled in, or opens neither a trip summary nor the search page, start from the search section's link with the same route, dates, trip type, and traveler counts. There is no link past the results, so selecting the flight and fare takes clicks.

### Dispatch snippet

```
If <CHECKOUT LINK> is an aircanada.com link, open it and dismiss any cookie banner. On a trip summary with flights <AC flight numbers> and <fare> already selected, skip to the trip review line. On an aircanada.com search form or results, stay there and skip the Open line. On anything else, also report checkout_link_failed with its URL.
Open https://www.aircanada.com/booking/us/en/aco/search?org0=<ORIGIN>&dest0=<DEST>&orgType0=A&destType0=A&departureDate0=<DD/MM/YYYY><RETURN LEG or omit>&adt=<ADULTS>&yth=<YOUTH>&chd=<CHILDREN>&inf=<LAP INFANTS>&ins=<INFANTS IN SEAT>&marketCode=<DOM|TBO|INT>&tripType=<OneWay|RoundTrip>&isFlexible=false
Wait up to 30 s for flight rows. If the summary's route, dates, or <ALL TRAVELERS> count differs, report it and stop.
If the <HH:MM> outbound row, its <AC flight numbers>, or its <cabin> price is missing, report fare_not_available and stop.
Select that row's <cabin> price, then <fare>; on an upgrade offer, press "Continue with <fare>". Pick the return row the same way.
On trip review, press Continue as guest. Press Continue on "Extend your session". After "Your session has timed out", press Resume last search and redo these steps with the travelers you have.
Close any "Family seating" dialog. With no seat to add, press Skip seats, then Continue, not Select seats. On Travel options, press Next, then Continue.
Extras, before or after the traveler page: add exactly <EXTRAS> and decline every other offer, upgrade, and insurance. If one won't add, report extra_unavailable with the page's reason and continue.
Stop at the first page that asks for traveler details, a sign-in, or payment. Report its URL, the fields it shows, the flights and fares selected, each extra's price, and the total shown, with one screenshot path.
If the total isn't for <ALL TRAVELERS - LAP INFANTS> seated travelers or a flight differs from the ones above, report the difference and stop. On Access Denied, a captcha, password, or 2FA prompt, stop and report the URL.
Ground rules for this Task. A Timeout or a net::ERR_* code from browser_navigate means the box cannot reach that host. Do not retry it, even if the tool result says to, and do not run DNS or connectivity checks. Move to the next source this Task names, or report the code and URL. An HTTP block (403, 419, 429, 503) follows this Task's own rule. Work in the tab you start in. Do not open a second tab, and switch tabs only when a tool result names a page you did not intend. In-page tabs are part of the page. A page script is an async (page) => function: run it with browser_run_code if you hold it. Without it, read rendered text from browser_snapshot or browser_find, read a <script> tag from the HTML that curl fetches in the box shell, and report a field neither yields as unreadable, never visible text in its place. Name the read you used. Snapshot a page state once. After a click, a scroll, or a wait this Task calls for, take one new snapshot. Do not snapshot the same state twice. Report as soon as the stop rule above is met, one line per item, only the fields asked for, no page dumps.
```

### When blocked

- Same tells as the search section.
- After a block, a site error, a stop code, or the user skipping a help request, do not dispatch again. Send the user the link the `flight-booking` skill's stopped-booking step names.
- If the page asks for a password, 2FA code, captcha, or payment, stop. Do not solve or bypass. Report which wall appeared and the URL it was on.

## Working with the subagent

Recipe steps name actions, not tools. A page script, a scroll, or a fill lands on whichever browser tools the subagent holds, and the ground rules say how, so paste the snippet as written. A subagent without a page-script tool reads what the snapshot or a shell fetch shows and names that read in its report.

Fill in every URL you know before you dispatch. A dispatch that changes something on the site (a cart, a booking, a form) never goes out without its URL. A read-only lookup may start from a name and a place when its section allows it. When the Fastest path is a shell call, run it with a short timeout and never queue a second lookup behind it. A failed call follows the section's own rule for that response (a single retry, a fresh key, or stop and report); when the section has no rule for it, or the call hangs, dispatch the subagent at once. When the section says the call is unreliable from the box, dispatch in the same round and stop the subagent with `StopSubagent` if the call returns first. A change to a job in flight goes to the same subagent, with `MessageSubagent` while it runs and `Task` with `resume` after it reports. A fresh dispatch spends rounds finding a tab and re-reading the page before it does anything.
