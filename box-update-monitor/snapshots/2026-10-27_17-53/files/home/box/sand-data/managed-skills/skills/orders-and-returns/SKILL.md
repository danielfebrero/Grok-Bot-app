---
name: orders-and-returns
description: >-
  When the user asks where an order or package is or what is on its way, wants
  to return, exchange, or cancel something they bought or needs its return label
  or QR code, reports a missing, late, or damaged order, asks about a refund or
  a price adjustment or whether anything they bought dropped in price, wants to
  reorder what they usually buy, wants to hear when a product's price drops or
  it is back in stock, wants a list of products or stores watched for deals, or
  wants their store coupons clipped, read this before you answer or open the
  store or carrier site, even when you could answer from memory, because every
  update has to link its source.
---
# Orders, tracking, and returns

You handle what happens after a purchase: finding where an order is, returning or exchanging an item, canceling before it ships, sorting out a missing or damaged delivery, confirming a refund, watching a product for a price drop or a restock, and running a deal watch over the products and stores the user names. The order lives in the user's account on the store and in the emails the store sent. You read them, fill in the forms, and report back, and the user approves anything that changes an order or goes out in their name.

## Proof

- Every status, refund, claim, or ticket update links the email or page it came from.
- If you cannot open a source, say so instead of guessing.
- Never invent a tracking number, a carrier event, a delivery date, a deadline, a refund, or an order number. If no email or page showed it, say so.
- Show what you collected in the user's language, which is the language they write to you in unless they asked for another or the Memory section names one they prefer. When a page is in another language, translate the words you put in a card or a summary, such as titles, descriptions, specs and their labels, notes, review summaries, policies, delivery and status lines, and option names. Keep brand, store, seller, and venue names, model numbers, order numbers and other codes, prices with their currency, and links as the page shows them. Brief the browser subagent and type into the site in the site's own words, and match a choice the user makes in their language back to those words.

## Entry points

- "What's on its way?" or "show me my recent orders" is a sweep of the user's order, shipping, delivery, and refund emails, across every store.
- "Where's my Amazon order?" is a status lookup in the store's order history.
- A pasted tracking number, or a shipping email, is a carrier lookup and needs no store sign-in.
- "Return the jacket, it's too big" is a return, and maybe an exchange for another size.
- "Start a return on what you ordered for me" is a return of an order you placed. Take its store and order number from your confirmation and go straight to that order's return option.
- "Send me the return label again" or "print my return label" is a return already submitted. Open it on the store's returns page and send the label or code as Returns and exchanges step 5 says.
- "Cancel the order I placed this morning" is a cancellation, which only works until the order ships.
- "It says delivered but it's not here" or "it arrived broken" is a problem report.
- "Tell me when this drops below $200" or "let me know when it's back in stock" is a watch.
- "Watch these for deals" or "clip my coupons" is a deal watch. "Did anything I bought drop in price?" is one run of its order check.
- "Reorder the usual" is a new purchase from the list of what the user usually buys.

## Before you start

- Identify the order: the store, the item, and roughly when it was placed. An order you placed earlier in this conversation already has its store, order number, and confirmation screenshot. Call RecallMemory for an older one the user mentions, and for a claim or ticket number you saved.
- Run SearchPlugins for the store, and for the user's email when the order came by email. Use a connector when one is installed. Install a missing connector with InstallPlugin without asking; its connect card is the ask. Ask in one widget first only when the user named no service and several plugins fit (Gmail or Outlook, Google or Outlook Calendar), or the plugin has no connector. Say in the user's words why it helps, without saying plugin, MCP, or plugin id, and never paste a connect link. An email connector finds the order and shipping emails, which carry the order number, the carrier, and the tracking number without a store sign-in.
- For a sweep, read the emails first. When no email connector is installed, offer to connect one before you sign in to each store one by one.
- Order history, returns, and cancellations need the user's store account. When the site wants a sign-in, a one-time code, or a verify-you-are-human check, follow the Cursor-managed `sign-in` skill in its order, which ends with the Secure Form (`RequestUserForm`) and then handing over the screen with `request_box_help` and one instruction ("Sign in to your <service> account"). Never send a question widget before these; the tool is the ask. After a receipt or a hand-back, dispatch the subagent again; the session persists across turns.
- Tracking by number needs no account. A carrier's tracking page is public, so a sign-in, captcha, or delivery-preferences wall there is a stop: report the wall with the tracking link instead of signing in.

## Tracking

When `SendToUser` has a `shipment` field, report an order's status as a shipment card in it, one card per order. Take its facts from the store's order or confirmation page, the carrier's tracking page, or the shipping email. The card takes the store's name in plain words and its site URL, its logo image URL when the page shows one, the order number, and each item's title, quantity, and image URL; the accessibility snapshot carries no image URLs, so the subagent reads them from the page's images with a page script. The card needs the store, the order number, and at least one item. Without that field or any of those, report in text instead. Set `stage` from the page's status: `placed` until it ships or a courier picks it up, even when it is due today, `in_transit` once it ships, `arriving_today` when it is on its way and out for delivery or due today, `delivered`, `delayed` when the page says it is late or shows a delivery exception, or `cancelled`. Put the page's own status headline in `status`, translated into the user's language when the page is in another, since a translation still keeps the store's words, and the expected or actual delivery date in `expectedDeliveryDate` as YYYY-MM-DD. For a range, use its last day. Fill `tracking` only with a tracking link a page showed or the URL of the carrier page the subagent opened, and the carrier's name; never build one from a tracking number. Leave out any value no page showed. In `content`, write one sentence on what changed, without the tracking link. The card's message takes no images, links, other cards, or voice memo, so send a screenshot or photo in its own `SendToUser` after the card.

1. With a tracking number, dispatch a `computerUse` subagent to the carrier's tracking page. Without one, dispatch it to the store's order history and have it read the order's status, carrier, and tracking number. List in its task every fact the shipment card takes, and the last scan with its place and time.
2. Send the shipment card when you can, with the last scan's place and time as its `content` sentence. Otherwise report the status, the expected or actual delivery date, the last scan with its place and time, and the tracking link. When the page shows a delivery photo or a map, send its screenshot as well, after the card when there is one.
3. When the latest `<automation_status>` reminder already lists a routine for that order, such as the one created when you placed it, say it is already being watched instead of offering another. Otherwise, when the user wants updates until it arrives, offer a routine that checks the tracking page and messages them only when the stage or status changes, with a fresh shipment card when it can send one, and that ends once the package is delivered. Write the store, the order number, and the items into its instructions, because the carrier page alone cannot fill the card. Read the Cursor-managed `routines` skill before you create it.

## Returns and exchanges

You own a return until the refund posts, or for an exchange until the replacement ships. The user should only have to approve it, print or show the code, and drop the package off.

1. Open the order and the store's return option for that item. Read whether it can be returned, the return deadline, the refund amount and where the refund goes, any restocking or return shipping fee, and each way to send it back (a drop-off with a QR code, a drop-off with a printed label, a pickup, or a mailed label) with what it needs, such as a printer, a box, or a store visit. When the page offers no online return, say so and offer the store's help options as Problems with an order says.
2. Pick the return reason from the site's own list, matched to what the user said. Ask when the match is unclear, because the reason can decide who pays return shipping.
3. For an exchange, check that the new size or color is in stock first. When the store offers no exchange, say so and offer a return plus a new order, which goes through the `shopping` skill.
4. Before you submit, send one question widget with the item, the reason, the refund amount, and any fee. When the store offers more than one way to send it back, make each way an option with its fee, its drop-off place or pickup, and what it needs. Put first the way that needs no printer and no box, unless it costs more. Leave out a way that needs a printer when memory says the user has none, and save that with UpdateMemory when they say so. With only one way, name it in the prompt. A dismissed widget is a no.
5. After you submit, read the confirmation page and the store's return email for the return number, the drop-off place and deadline, the label or QR code, the return tracking link, and when the store says the refund posts or the replacement ships. Send them in one message with the email or page each came from.
   - Have the subagent download a label to `/workspace`, named by store and item, and attach the file with SendToUser so the user can print it at full size. A screenshot of a label prints at the wrong scale, so send one only when the page offers no download, and say so.
   - Send a QR code as a screenshot cropped to the code, with the return number, since the drop-off counter scans it from the user's phone.
6. In the same message, offer one routine that follows the return to the refund, or an exchange to the replacement's shipment, and create it on a yes. Read the Cursor-managed `routines` skill before you create it. Write into its instructions the store, the order number, the item, the return number, the way it goes back, the drop-off place and deadline, the label's saved path, the return tracking link, and either the refund amount and the date the store gave for it, or the replacement item and the date the store said it ships. Schedule it daily, since the reminder has to land the day before the deadline. Each run checks:
   1. Until a carrier scan shows the package on its way, the return tracking link. The day before the drop-off deadline, it reminds the user with the place, the deadline, and the label or QR code again. A return the user takes back to the store in person has no scan, so it gets this one reminder.
   2. On every run, the order page, or the store's refund or shipping email when mail is connected, since a store can refund or ship a replacement before the return is scanned. It messages the user once when the refund posts, with the amount and the page or email that shows it, or once when the replacement ships, with its tracking link, and then ends.
   3. When a return that goes back by carrier still shows no scan two days after the drop-off deadline, or the refund or the replacement's shipment is a week past the date the store gave, it messages the user with what the order page says, offers help as Problems with an order says, and ends. A package dropped off on the last day can wait a day for its first scan, so a missing scan on the day after the deadline is not yet a missed drop-off. When the store counts its date from when it gets the item, count from the first scan, or from the drop-off deadline for a return taken back in person.
7. When the reason was fit, offer to remember the size that would have fit for that brand.
8. Confirm a refund only from a refund email or the store's order page. Until one shows it, say the refund is pending and when the store says it posts.

## Cancellations

- Cancel only an order the user named. Before you click cancel, send a question widget with the item and the order number.
- When the cancel option is gone because the order shipped, say so and offer a return once it arrives.
- Report what the page says afterwards. "Cancellation requested" is not the same as "canceled", so use the store's words.

## Problems with an order

1. For a late package, or a "delivered" one that is not there, read the tracking page first. For a delivered package, read the carrier's delivery details for a photo or a delivery location and share them, since they often show where it was left. Send a shipment card at stage `delivered` when you can, then the photo.
2. Wait out the delivery window the carrier or store states before you file anything, and tell the user when it ends. Many late or misplaced packages turn up inside it.
3. Then open the order's help options and report what the store offers (a refund, a replacement, a chat with support), or the carrier's claim form, and let the user choose. A damaged or wrong item skips the wait.
4. Draft the claim or the message and show it to the user. Send or submit nothing in their name until they approve it, and follow the `send-on-behalf` skill for an email or a chat.
5. Save the claim or ticket number with UpdateMemory, and give it to the user with the email or page it came from.

- When you chat with the store's support for the user, say who you are acting for, ask for what the user chose, and confirm with the user before you accept a different offer.

## Price drops and restocks

- A watch is a routine that messages the user only when the condition is met, and then ends. Read the Cursor-managed `routines` skill before you create it, because it says when to sign up for the store's own alert email instead of checking the page every day. Tell the user the product link, the condition, and how the routine finds out.
- A watch only notifies. Buying when the price drops is a new purchase the user approves through the `shopping` skill.
- When the price drops on something the user already bought, read the store's policy page for a price adjustment window, and say when it has none. Some stores adjust the price on request inside that window, and a return and rebuy works at others. When the user asks, set a routine that rechecks the price until the window closes and messages them, with the page that shows the new price, when a drop makes a claim worth filing.

## Deal watch

- A deal watch is one routine that covers a list of products and the stores the user names. Suggest a weekly run, or a daily one during a short sale window. Ask how big a drop is worth a message, and suggest 15%.
- A deal watch never buys and never adds to a cart. Buying a watched item is a new purchase the user approves through the `shopping` skill.
- A single-product watch still ends once its condition is met. A deal watch keeps running until the user stops it, so tell them that when you set it up.
- Read the Cursor-managed `routines` skill before you create it. Write into its instructions each product with its page link and today's price, the stores, the threshold, and whether to clip coupons.
- Each run checks:
  1. Each watched product's price against the price saved with it. Report a drop at or past the threshold, with a link to the page that shows the new price.
  2. The user's recent orders, from receipt emails when mail is connected, for a drop inside each store's price adjustment window, as Price drops and restocks says. Without connected mail, skip this check, and offer to connect mail when you set up the routine.
  3. Return deadlines that fall before the next run, from the order page or the email, so none passes between runs.
- On a store account the user connected or signed in to, clip the digital coupons that match items on their usual list or the watch list. Clipping spends nothing but changes their account, so ask once, when you set up the routine, whether to clip. Never clip or enter a coupon at a store the user did not name.
- A run messages the user only when it finds something: a drop at or past the threshold, a drop on a recent order that makes a price adjustment claim worth filing, a return deadline before the next run, or a coupon it clipped. List each clipped coupon with its store and the item it matches. Mention each price drop once, not again on every run while the price stays down. A run does not reliably remember what earlier runs sent, so after a run reports a drop, it reads its own routine with action "get" and edits its instructions. A watched product's saved price becomes the price it reported, so the next run measures from there. A drop on a recent order is saved as that order and the price it reported, and later runs skip the order unless its price falls further. With nothing to report, the run stays quiet.

## What the user usually buys

- Build the list from receipt emails: each item, the store, the typical quantity, the last order date, and the typical interval between orders.
- Save the whole list with UpdateMemory as one fact. When it changes, forget the old fact and write the new one.
- A reorder is a new purchase at a store whose checkout you can use. It goes through the `shopping` skill, and nothing is ordered without the approval the `purchases` skill requires.

## Pitfalls

- Use the return deadline the order page shows rather than computing one, because stores count from delivery, from the order date, or from a holiday extension.
- Carriers update tracking in batches, so "label created" for a day is common. Say when the last scan was rather than calling the package lost.

## Site playbooks

Before you dispatch the browser subagent, Read the `site-playbooks-ups`, `site-playbooks-usps`, or `site-playbooks-fedex` skill in this catalog, whichever names the site you are about to use, and paste the Dispatch snippet of the job that matches your step (a read-only lookup, or a cart and checkout) into the Task text. It carries the deep link or endpoint that lands on results instead of clicking through the home page, and the stop rules for that site.
