---
name: site-playbooks-google-flights
description: >-
  Google Flights site playbook: cheapest fares on a route, prices across a date
  window.
---
# Site playbook: Google Flights

Read this before you dispatch a computerUse subagent to this site. Follow the Fastest path, fill in every URL you know, then paste the Dispatch snippet into the Task text as written. The subagent does not read skills, so every rule it needs is in the snippet, including the ground rules at its end. Secrets never go in the Task text. They come through `RequestUserForm`.

Jobs in this file: Google Flights cheapest itinerary search; Google Flights prices across a date window.

## Google Flights cheapest itinerary search

### When to use

The user wants the cheapest one way or round trip fares between two airports on fixed dates, without booking. The parent needs origin and destination IATA codes, the depart date, the return date for round trips, and every traveler, with children's ages. The user gets the lowest priced itineraries with price, airline, total duration, stops, depart and arrive times, and the results URL as the booking link. When `FlightSearch` is in your tools, use this only when it finds no fares or the user asks for Google Flights. The results URL is also the link to send when a booking can't finish in the browser; build it without dispatching a subagent.

### Fastest path

Build the `tfs` parameter with the Node script below in the box shell, then `browser_navigate` the results URL. No form filling, no clicks. Build one `search?tfs=` URL per route and date pair in the shell, then paste the finished URL into every computerUse Task, including re-dispatches and resumes. Never send the subagent to the Flights home page to type airports and dates. `tfs` is a protobuf, base64url encoded (`+` to `-`, `/` to `_`, strip `=` padding), with fields in this wire order.

| Field | Wire type | Value |
|---|---|---|
| `f1`, `f2` | varint | `28`, `2` |
| `f3` | message, repeated, one per leg | `{ f2: "YYYY-MM-DD" (date), f13: {f1:1, f2:"<ORIGIN>"}, f14: {f1:1, f2:"<DEST>"} }` |
| `f8` | varint, repeated, one per traveler | `1`=adult, `2`=child, `3`=infant on lap; a child 12 or older is an adult |
| `f9` | varint | cabin, `1`=economy, `2`=premium economy, `3`=business, `4`=first |
| `f14`, `f16` | varint, message | `1`, `{ f1: 2^64 - 1 }` (no price cap) |
| `f19` | varint | trip type, `1`=round trip, `2`=one way |

One way is exactly one `f3` leg with `f19=2`. Round trip is two `f3` legs (the second reverses airports and uses the return date) with `f19=1`. `f19` is the trip type field, not `f2`. A single leg with `f19=1` or no `f19` renders round trip fares labeled "round trip". Encode every traveler, so the link the user gets opens with their party.

```js
// node tfs.js in the box shell
const PAX = {"adult":1,"child":2,"infant_on_lap":3}, CABIN = {"economy":1,"premium_economy":2,"business":3,"first":4};
function v(n){const o=[];while(n>127){o.push((n%128)|0x80);n=Math.floor(n/128);}o.push(n);return Buffer.from(o);}
function vf(f,n){return Buffer.concat([v(f*8),v(n)]);}
function mf(f,b){return Buffer.concat([v(f*8+2),v(b.length),b]);}
const sf = (f,s) => mf(f, Buffer.from(s)), airport = c => Buffer.concat([vf(1,1), sf(2,c)]);
const leg = (date,from,to) => Buffer.concat([sf(2,date), mf(13,airport(from)), mf(14,airport(to))]);
function tfs({oneway, origin, dest, depart, ret, cabin = "economy", pax = ["adult"]}) {
  const p = [vf(1,28), vf(2,2), mf(3, leg(depart, origin, dest))];
  if (!oneway) p.push(mf(3, leg(ret, dest, origin)));
  for (const t of pax) p.push(vf(8, PAX[t]));
  p.push(vf(9, CABIN[cabin]), vf(14,1), mf(16, Buffer.from([8,255,255,255,255,255,255,255,255,255,1])), vf(19, oneway ? 2 : 1));
  return Buffer.concat(p).toString("base64url");
}
// tfs({"oneway":true,"origin":"SFO","dest":"JFK","depart":"2026-07-15"}) gives "CBwQAhoeEgoyMDI2LTA3LTE1agcIARIDU0ZPcgcIARIDSkZLQAFIAXABggELCP___________wGYAQI"
// tfs({"origin":"SFO","dest":"JFK","depart":"2026-07-15","ret":"2026-07-22","pax":["adult","adult","child","infant_on_lap"]}) gives "CBwQAhoeEgoyMDI2LTA3LTE1agcIARIDU0ZPcgcIARIDSkZLGh4SCjIwMjYtMDctMjJqBwgBEgNKRktyBwgBEgNTRk9AAUABQAJAA0gBcAGCAQsI____________AZgBAQ"
```

Results URL. `&curr=USD` forces USD and `&hl=en` forces English. Always append both so the box locale does not change currency or language.

```
https://www.google.com/travel/flights/search?tfs=<TFS>&curr=USD&hl=en
https://www.google.com/travel/flights/search?tfs=CBwQAhoeEgoyMDI2LTA3LTE1agcIARIDU0ZPcgcIARIDSkZLQAFIAXABggELCP___________wGYAQI&curr=USD&hl=en
```

### Dispatch snippet

```
Open https://www.google.com/travel/flights/search?tfs=<TFS>&curr=USD&hl=en (<TFS> encodes <ORIGIN> to <DEST>, depart <YYYY-MM-DD>, return <YYYY-MM-DD or none>, and every traveler).
Wait about 4 s after load. Confirm the title reads "<Origin city> to <Destination city> | Google Flights".
Run the page script async (page) => page.locator('[role=main]').first().innerText() and parse every row: depart time, arrive time (+1 = next day), airline, total duration, stops, route, $price.
Skip rows without an airline or airport codes, dedupe, sort by price ascending, and return the cheapest <N> as a table with the results URL and one screenshot path.
Read only. Do not click Select flight. If a captcha, sign-in, or payment wall appears, stop and report the URL.
Ground rules for this Task. A Timeout or a net::ERR_* code from browser_navigate means the box cannot reach that host. Do not retry it, even if the tool result says to, and do not run DNS or connectivity checks. Move to the next source this Task names, or report the code and URL. An HTTP block (403, 419, 429, 503) follows this Task's own rule. Work in the tab you start in. Do not open a second tab, and switch tabs only when a tool result names a page you did not intend. In-page tabs are part of the page. A page script is an async (page) => function: run it with browser_run_code if you hold it. Without it, read rendered text from browser_snapshot or browser_find, read a <script> tag from the HTML that curl fetches in the box shell, and report a field neither yields as unreadable, never visible text in its place. Name the read you used. Snapshot a page state once. After a click, a scroll, or a wait this Task calls for, take one new snapshot. Do not snapshot the same state twice. Report as soon as the stop rule above is met, one line per item, only the fields asked for, no page dumps.
```

### Read the result

```js
async (page) => page.locator('[role=main]').first().innerText()   // all rows, about 3 KB
async (page) => page.locator('li.pIav2d').allInnerTexts()         // one string per row
```

Each row reads like `2:35 PM` dash `11:20 PM JetBlue 5 hr 45 min SFO` dash `JFK Nonstop 414 kg CO2e ... $205`, where each dash is an en dash (U+2013) on the page. Multi stop rows add `1 stop` and a connection like `50 min PHX`.

Parse per row. Depart time, arrive time (a trailing `+1` means next day arrival), airline, total duration, stops (`Nonstop` is 0, `N stop` or `N stops` is N), route as two IATA codes joined by U+2013, price `$NNN`.

For a group, quote the total only from FlightSearch or the airline's checkout.

Rows are duplicated in the DOM, once fully (airline plus airport codes) and once condensed. Skip rows lacking an airline name or airport codes, then dedupe by (depart, arrive, airline, price). The list defaults to Google's "Best" ranking, so sort ascending by price yourself. Flight numbers are not in the list view; they appear only after expanding a row's chevron ("Operated by ... as ... flight NNNN"), which is read only but costs turns. Times are local to each airport and durations already account for the offset. Round trip prices are the total fare.

### When blocked

- In earlier testing no captcha or 403 was hit. Proxies were used only to keep currency and locale stable, which `&curr=USD&hl=en` handles. Not yet verified on a box.
- Zero priced rows in `[role=main]` under the right title means no flights for the route or date. Report it as a real empty result with the URL.
- If the title does not name the requested cities, the page is not the route you asked for.
- Recheck the IATA codes and rebuild `tfs`.
- There is no usable JSON API. The internal `batchexecute` and GRPC endpoints are obfuscated, so do not hunt for one.
- If the page asks for a password, 2FA code, captcha, or payment, stop. Do not solve or bypass. Report which wall appeared and the URL it was on.

## Google Flights prices across a date window

### When to use

The user's window allows more than 4 date pairs, so one fare search can't cover it. The parent needs origin and destination IATA codes, the first and last dates of the window, at most 3 trip lengths, the cabin, every traveler with children's ages, and the stop limit. It gets back a price for each date pair, unsorted, and picks the pairs to search itself. For a metro, read the calendar for its main airport, then search fares from every airport in it: a metro `origin` in `FlightSearch` covers them in one call, and a site search needs one search per airport.

### Fastest path

Build `tfs` with the script in Google Flights cheapest itinerary search: depart on the window's first day and, for a round trip, return one trip length later. Paste the finished URL into the computerUse Task.

The departure date picker shows two months of days, each with the lowest round trip price for the trip length currently set, or the one way price. Read it, not the Date grid: the grid shows only 7 departure days around the current dates, and its cells mix trip lengths. Set the stop limit with the Stops filter before reading.

### Dispatch snippet

```
Open https://www.google.com/travel/flights/search?tfs=<TFS>&curr=USD&hl=en (<TFS> encodes <ORIGIN> to <DEST>, depart <FIRST DAY OF WINDOW>, return <DEPART + TRIP LENGTH, or none for one way>, the cabin, and every traveler).
Wait about 4 s after load. Confirm the title reads "<Origin city> to <Destination city> | Google Flights".
<If there is a stop limit:> Open Stops and pick <Nonstop only | 1 stop or fewer>. If you can't set it, say the prices below ignore the stop limit and are rough.
Click the departure date field. Read every day from <FIRST DAY> to <LAST DAY> with its price from one snapshot, using the next month arrow when the window runs past the months shown.
<For each other trip length:> Set the return date that many days after the departure, then read the days again.
Return one line per date pair: depart date, return date or none, and price, or "no price" when a day shows none. Do not sort or pick. Add the results URL, whether the stop limit was set, and one screenshot path.
Read only. Do not click Select flight. If a captcha, sign-in, or payment wall appears, stop and report the URL.
Ground rules for this Task. A Timeout or a net::ERR_* code from browser_navigate means the box cannot reach that host. Do not retry it, even if the tool result says to, and do not run DNS or connectivity checks. Move to the next source this Task names, or report the code and URL. An HTTP block (403, 419, 429, 503) follows this Task's own rule. Work in the tab you start in. Do not open a second tab, and switch tabs only when a tool result names a page you did not intend. In-page tabs are part of the page. A page script is an async (page) => function: run it with browser_run_code if you hold it. Without it, read rendered text from browser_snapshot or browser_find, read a <script> tag from the HTML that curl fetches in the box shell, and report a field neither yields as unreadable, never visible text in its place. Name the read you used. Snapshot a page state once. After a click, a scroll, or a wait this Task calls for, take one new snapshot. Do not snapshot the same state twice. Report as soon as the stop rule above is met, one line per item, only the fields asked for, no page dumps.
```

### Read the result

Calendar prices are Google's estimates for the trip length set, not fares. Report them only as a way to choose dates; the fares come from the search that follows.

### When blocked

- The date picker read is not yet verified on a box.
- A day without a price means Google has no estimate for it. Report it as no price, never as no flights.
- If the picker shows no prices at all, report that with the URL and one screenshot path.
- If the page asks for a password, 2FA code, captcha, or payment, stop. Do not solve or bypass. Report which wall appeared and the URL it was on.

## Working with the subagent

Recipe steps name actions, not tools. A page script, a scroll, or a fill lands on whichever browser tools the subagent holds, and the ground rules say how, so paste the snippet as written. A subagent without a page-script tool reads what the snapshot or a shell fetch shows and names that read in its report.

Fill in every URL you know before you dispatch. A dispatch that changes something on the site (a cart, a booking, a form) never goes out without its URL. A read-only lookup may start from a name and a place when its section allows it. When the Fastest path is a shell call, run it with a short timeout and never queue a second lookup behind it. A failed call follows the section's own rule for that response (a single retry, a fresh key, or stop and report); when the section has no rule for it, or the call hangs, dispatch the subagent at once. When the section says the call is unreliable from the box, dispatch in the same round and stop the subagent with `StopSubagent` if the call returns first. A change to a job in flight goes to the same subagent, with `MessageSubagent` while it runs and `Task` with `resume` after it reports. A fresh dispatch spends rounds finding a tab and re-reading the page before it does anything.
