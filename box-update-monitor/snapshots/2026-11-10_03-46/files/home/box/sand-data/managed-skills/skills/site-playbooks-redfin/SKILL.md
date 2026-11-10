---
name: site-playbooks-redfin
description: >-
  Redfin site playbook: home search or one listing; not tours, offers, or agent
  contact.
---
# Site playbook: Redfin

Read this before you dispatch a computerUse subagent to this site. Follow the Fastest path, fill in every URL you know, then paste the Dispatch snippet into the Task text as written. The subagent does not read skills, so every rule it needs is in the snippet, including the ground rules at its end. Secrets never go in the Task text. They come through `RequestUserForm`.

Jobs in this file: Redfin home search; Redfin listing page.

## Redfin home search

### When to use

The user wants Redfin homes for sale, rentals, or recent sales in a city, ZIP code, or neighborhood, filtered by price, beds, baths, or home type (house or condo), or cheapest first. The parent needs the area (a ZIP, or a city or neighborhood with its state), the listing type, and the filters. The user gets each home's address, price, beds, baths, square feet, status labels, and Redfin link, plus the area's total. Read only.

### Fastest path

`browser_navigate` an area URL with the filters in one path segment, then read the cards with the page script; nothing is typed in the search box. Area URLs are `https://www.redfin.com/zipcode/<zip>`, `/city/<id>/<ST>/<City>`, and `/neighborhood/<id>/<ST>/<City>/<Name>`. A bare `/<ST>/<City>` is a 404 Page Not Found.

Ids come from links. A ZIP page links each nearby city as "<City> homes for sale" with its `/city/<id>/<ST>/<City>` path, and a city page links its neighborhoods ("<Name> homes for sale", `/neighborhood/<id>/...`) and its ZIP codes.

Put a ZIP you know in the user's city into the dispatch, so the subagent's first page is that ZIP page. Put a `/city/` or `/neighborhood/` path a report returns into later dispatches for the same place.

| Part | Form |
|---|---|
| For rent | `/rentals` after the area path; an `/apartments-for-rent` link redirects there |
| Recently sold | `include=sold-3mo` in the filter segment |
| Filters | one `/filter/` segment after the area or `/rentals`, comma joined: `min-price=400k`, `max-price=900k` (prices in thousands with a `k`, a month for rentals), `min-beds=2`, `min-baths=2`, `property-type=house` or `condo` |
| Sort | `sort=lo-price` in the same segment, cheapest first |
| Page | `/page-<n>` after the filter segment; the pager's own links drop the filters |

Worked example. `https://www.redfin.com/city/30818/TX/Austin/filter/sort=lo-price,property-type=house,max-price=600k,min-beds=3,min-baths=2/page-2`

Redfin may reorder the filters in the URL, which changes nothing. The heading still says "homes for sale" under a sold filter, so the Filters button is the check that filters took: it reads "Filters • N", counting each filter kind once, a min and max price as one, and not the sort.

`redfinSearchCards` reads the `.bp-Homecard[title]` cards in each `[data-rf-test-id="photos-view"]` group and skips cards with no title. It returns `count`, `page`, `nearby[]` from any group after the "Nearby homes that match your criteria" heading (matched on its first four words), and `matches[]` from the rest, each card with its `address` (the `title`), `href`, `sashes`, and `text`: price, beds, baths, square feet, address, and a features line.

Each card also has a JSON-LD block (`SingleFamilyResidence` plus `Product`) with name, url, address, geo, `numberOfRooms`, `floorSize`, and `offers.price`, but no baths or status, so the cards are the fuller read.

`robots.txt` disallows `/stingray/`, so the JSON endpoints under it are off limits. Read the pages.

### Dispatch snippet

```
Find Redfin homes <for sale|for rent|sold> in <area> with <filters>. Read only. Never press Send message, the heart, Save search, or Sign in.
With no area path, take the "<City> homes for sale" href on https://www.redfin.com/zipcode/<ZIP in the city>, and for a neighborhood the "<Name> homes for sale" href on that city page.
Open https://www.redfin.com<path></rentals if for rent>/filter/<comma-joined filters> and wait 3 s (no filters: drop /filter/).
If "Filters • N" counts fewer filters than sent (a price range once, sort not), report filter_not_applied with what it shows and stop.
Run the page script async (page) => { const r = { matches: [], nearby: [] }; for (const v of await page.locator('[data-rf-test-id="photos-view"]').all()) { const k = (await v.locator('xpath=preceding::*[not(self::script)][contains(text(), "Nearby homes that match")]').count()) ? 'nearby' : 'matches'; for (const h of await v.locator('.bp-Homecard[title]').all()) r[k].push({ address: await h.getAttribute('title'), href: await h.locator('a[href]').first().getAttribute('href', { timeout: 2e3 }).catch(() => null), sashes: (await h.locator('[data-rf-test-id="home-sash"]').allInnerTexts()).filter((s) => !/about this home/i.test(s)), text: (await h.locator('.bp-Homecard__Content').allInnerTexts()).join('\n') }); } const t = async (l) => (await l.allInnerTexts())[0]?.trim() ?? null; return { count: await t(page.locator('.homes.summary')), page: await t(page.getByText(/^Viewing page/)), ...r }; }
Report count, page, path, and per matches[] card address, price, beds, baths, sq ft, sashes, https://www.redfin.com<href>; nearby[] apart, as outside the area.
More pages: append /page-<n> after the filters, up to <PAGES>.
On "Let's confirm you are human" (a Begin button), a captcha, or Access Denied, report redfin_wall with the URL and stop; never press Begin. On a sign-in, 2FA, or payment page, stop and report the URL. Take one screenshot.
Ground rules for this Task. A Timeout or a net::ERR_* code from browser_navigate means the box cannot reach that host. Do not retry it, even if the tool result says to, and do not run DNS or connectivity checks. Move to the next source this Task names, or report the code and URL. An HTTP block (403, 419, 429, 503) follows this Task's own rule. Work in the tab you start in. Do not open a second tab, and switch tabs only when a tool result names a page you did not intend. In-page tabs are part of the page. A page script is an async (page) => function: run it with browser_run_code if you hold it. Without it, read rendered text from browser_snapshot or browser_find, read a <script> tag from the HTML that curl fetches in the box shell, and report a field neither yields as unreadable, never visible text in its place. Name the read you used. Snapshot a page state once. After a click, a scroll, or a wait this Task calls for, take one new snapshot. Do not snapshot the same state twice. Report as soon as the stop rule above is met, one line per item, only the fields asked for, no page dumps.
```

### Read the result

- About 40 cards load per page without scrolling, under "Viewing page X of Y". The header counts the area's matches as "N homes", sometimes with "+ M early access", or "N rentals". When an area has few matches, "End of results" and "Nearby homes that match your criteria" follow, with homes outside the area that the count leaves out; those are `nearby[]`.
- Sashes are labels such as NEW <time> AGO, OPEN <day and hours>, EARLY ACCESS, FOR SALE BY OWNER, and SOLD <date>. A sold card's price is the sale price. A rental card shows `$<rent>/mo`. An apartment building card shows its lowest rent with `+` and "Total price", ranges for beds, baths, and square feet, "N units match", a phone number, and Send message, and links to `/apartment/<id>` or `/home/<id>`.

### When blocked

- Redfin's wall is an AWS WAF page titled "Let's confirm you are human" with an orange Begin button, served as HTTP 405 with `x-amzn-waf-action: captcha` and scripts from `captcha.awswaf.com`. It can cover every page, the home page too. A plain fetch gets the same 405 with a page titled "Human Verification".
- After `redfin_wall`, do not dispatch to Redfin again this turn. Hand the user the filtered Redfin URL to open themselves, and move to the Realtor.com playbook unless Realtor.com already blocked this turn, then to the county assessor or GIS portal its When blocked describes. After `filter_not_applied`, check the filter names against the Fastest path and dispatch once more.
- This page access was checked only in a desktop browser outside Grok Bot's production network, after a person cleared Redfin's check, not from the box. Plan that fallback before the first dispatch.
- If the page asks for a password, 2FA code, captcha, or payment, stop. Do not solve or bypass. Report which wall appeared and the URL it was on.

## Redfin listing page

### When to use

The user has a Redfin home link (`/<ST>/<City>/<address>/home/<id>`) or picked a home from the search section, and wants its price, status, Redfin Estimate, facts, HOA dues, property tax and sale history, days on market, or listing agent and brokerage. The parent needs the link; for an address alone, run the search section on its ZIP and take the matching card's link. Read only.

### Fastest path

`browser_navigate` the link. The page title reads "<address> | MLS# <n> | Redfin" (some sold homes have no MLS#), and every section is in the page on load, so no scrolling is needed.

`redfinListing` reads the `[data-rf-test-id]` sections `sashContainer` (status), `abp-homeinfo-homemainstats` (price, beds, baths, square feet), `abp-homeinfo-homeaddress`, and `house-info` (remarks, key details, Listed by, days on Redfin), and the Redfin Estimate from `avm-section-expandable-preview` on an active listing or `#redfin-estimate-scroll` on a sold one. It presses Enter on the Sale History tab, then on the Tax History tab, which a sticky header covers from clicks, and reads the tab panel in the same section as each tab.

The page also has a `RealEstateListing` JSON-LD block with `offers.price`, `offers.availability`, and the year built, bedrooms, bathrooms, and floor size under `mainEntity`.

### Dispatch snippet

```
Read the Redfin listing <listing URL>. Read only. Never press Request showing, Start an offer, Ask a question, Contact, Get prequalified, Favorite, I own this home, Edit Facts, or Sign in.
Open the URL, wait 3 s, and run the page script async (page) => { const t = async (s) => (await page.locator(s).allInnerTexts())[0]?.trim() ?? null, i = (id) => '[data-rf-test-id="' + id + '"]', h = async (n, f) => { const b = page.getByRole('tab', { name: n, exact: true }).first(); if (!(await b.count())) return null; await b.press('Enter'); return b.locator('xpath=ancestor::*[.//*[@role="tabpanel"]][1]').getByRole('tabpanel').filter(f).first().innerText({ timeout: 3e3 }).catch(() => null); }; return { status: await t(i('sashContainer')), stats: await t(i('abp-homeinfo-homemainstats')), address: await t(i('abp-homeinfo-homeaddress')), about: await t(i('house-info')), estimate: ((await t(i('avm-section-expandable-preview'))) ?? (await t('#redfin-estimate-scroll')))?.slice(0, 300) ?? null, saleHistory: await h('Sale History', { hasNotText: 'Property tax' }), taxHistory: await h('Tax History', { hasText: 'Property tax' }) }; }
Report status (the sash; empty is active), price, beds, baths, sq ft, address; from about the key details, "Listed by <agent> • <brokerage>", "Bought with" if sold, and "<N> days on Redfin", else the latest Listed date in saleHistory.
Report the Redfin Estimate and its gap to list price, or no_estimate if estimate is null or says it cannot show one; each saleHistory row as date, event, price; the latest three taxHistory rows as year, property tax, assessment; the MLS# from the page title.
Never click Unlock MLS, the payment calculator, or tour dates.
On "Let's confirm you are human" (a Begin button), a captcha, or Access Denied, report redfin_wall with the URL and stop; never press Begin. On a sign-in, 2FA, or payment page, stop and report the URL. Take one screenshot.
Ground rules for this Task. A Timeout or a net::ERR_* code from browser_navigate means the box cannot reach that host. Do not retry it, even if the tool result says to, and do not run DNS or connectivity checks. Move to the next source this Task names, or report the code and URL. An HTTP block (403, 419, 429, 503) follows this Task's own rule. Work in the tab you start in. Do not open a second tab, and switch tabs only when a tool result names a page you did not intend. In-page tabs are part of the page. A page script is an async (page) => function: run it with browser_run_code if you hold it. Without it, read rendered text from browser_snapshot or browser_find, read a <script> tag from the HTML that curl fetches in the box shell, and report a field neither yields as unreadable, never visible text in its place. Name the read you used. Snapshot a page state once. After a click, a scroll, or a wait this Task calls for, take one new snapshot. Do not snapshot the same state twice. Report as soon as the stop rule above is met, one line per item, only the fields asked for, no page dumps.
```

### Read the result

- The price line pairs the price with "Est. $X/mo", a mortgage payment estimate, not the Redfin Estimate. The estimate section reads "Redfin Estimate" with the amount and how far it sits under or over the list price. On some active listings it says "We cannot show automated home-value estimates for this home" and lists nearby sales; those are comparables, not this home's value.
- Key details show each value above its label: Property Type, Year Built, Lot Size, Price/Sq.Ft., HOA Dues when the home has dues, and Parking. Then come "Listed by <agent> • <brokerage>", "Bought with" on a sold home, "Listing updated", and "N days on Redfin • N views • N favorites", which some listings show without the days.
- Sale History rows read date, event (Listed, Listed (Active), Price Changed, Contingent, Pending, Sold), source (an MLS number or Public Records), price, and price per square foot, with Show more at the end of a long history. Where a state does not disclose sale prices, a sold row shows `*` for its price. An "Unlock MLS #<n>" line above the rows is a link, not a row. Tax History rows read year, property tax with its yearly change, land plus additions, and assessment.

### When blocked

- Same wall tells as the search section.
- After `redfin_wall`, do not dispatch to Redfin again; hand the user the link. Try the Realtor.com playbook's single-address branch first, unless Realtor.com already blocked this turn, then the county assessor or GIS portal its When blocked describes. A `no_estimate` report is the page's own answer, not a failed read.
- If the page asks for a password, 2FA code, captcha, or payment, stop. Do not solve or bypass. Report which wall appeared and the URL it was on.

## Working with the subagent

Recipe steps name actions, not tools. A page script, a scroll, or a fill lands on whichever browser tools the subagent holds, and the ground rules say how, so paste the snippet as written. A subagent without a page-script tool reads what the snapshot or a shell fetch shows and names that read in its report.

Fill in every URL you know before you dispatch. A dispatch that changes something on the site (a cart, a booking, a form) never goes out without its URL. A read-only lookup may start from a name and a place when its section allows it. When the Fastest path is a shell call, run it with a short timeout and never queue a second lookup behind it. A failed call follows the section's own rule for that response (a single retry, a fresh key, or stop and report); when the section has no rule for it, or the call hangs, dispatch the subagent at once. When the section says the call is unreliable from the box, dispatch in the same round and stop the subagent with `StopSubagent` if the call returns first. A change to a job in flight goes to the same subagent, with `MessageSubagent` while it runs and `Task` with `resume` after it reports. A fresh dispatch spends rounds finding a tab and re-reading the page before it does anything.
