---
name: subscription-management
description: >-
  When the user asks to cancel a subscription they named, to read the price and
  renewal date on a service they named, to start a free trial and leave before
  it bills, to find paid subscriptions they no longer use or pay for twice, to
  watch a subscription price for a change, or to ask a service to refund a
  subscription charge, read this before you open the billing or signup page, or
  their receipt mail, even when you could answer from memory, because the price,
  the date, and proof of cancellation have to come from a page or email opened
  this turn.
---
# Subscriptions

You cancel a subscription the user named, or you read its price and renewal date. Later sections add other jobs. A job that does not have its own section below is not specified yet. Say that, and do not invent the missing steps.

## Shared rules

These rules apply to every job in this skill.

- The service and the plan come from the user. Ask only when they named neither.
- Run SearchPlugins for the service, and for the user's email. Use a connector when one is installed. Install a missing connector with InstallPlugin without asking; its connect card is the ask. Ask in one widget first only when the user named no service and several plugins fit (Gmail or Outlook, Google or Outlook Calendar), or the plugin has no connector. Say in the user's words why it helps, without saying plugin, MCP, or plugin id, and never paste a connect link. An email connector finds the signup confirmation and the cancellation message.
- When the site wants a sign-in, a one-time code, or a verify-you-are-human check, follow the Cursor-managed `sign-in` skill in its order, which ends with the Secure Form (`RequestUserForm`) and then handing over the screen with `request_box_help` and one instruction ("Sign in to your <service> account"). Never send a question widget before these; the tool is the ask. After a receipt or a hand-back, dispatch the subagent again; the session persists across turns.
- Every price, date, and status comes from a page or email opened for this task. When you cannot open one, say so instead of guessing.
- A signup, a plan change, or a cancellation starts with one question widget that states the service, the plan, the price, and the date. A dismissed widget is a no. A yes opens that page. It does not answer a retention offer, and it does not answer a final confirm. Each later submit is its own `computerUse` dispatch, and it is not read-only.
- A retention offer, such as a cheaper plan, a pause, extra free time, or a discount, stops the job. Send its own question widget, with the renewal price and the end date, and end the turn. Do not click Continue, Confirm, or any control that leaves that page until the user answers that widget. The user's request to cancel is not that answer, and neither is a yes on the earlier question. Taking the offer stops the cancellation. Declining the offer, or dismissing that widget, is the only way to go on.
- The final confirm page stops the job again. Send a question widget that names the service, the plan, and the date access ends, and end the turn. Do not click the confirm control until that widget's yes. The user's request to cancel, a yes on the first question, and a decline of a retention offer are not that yes. A later section does not waive either stop.
- Report a subscription as cancelled only when a confirmation page or the service's email states that it is cancelled and names the date access ends. Send that page or email, with a screenshot. Save the service, the plan, the end date, and the link with UpdateMemory.
- When the page says the cancellation is requested, pending, or processing, use those words and say it is not confirmed yet.
- When no page confirms the cancellation, say that in the same message. Give the phone number, the hours, and the menu path the page shows. When place_phone_call is among your tools, offer to place that call, and place it only after the user says yes. Report it cancelled only after the call's result or a later page or email states the cancellation and the end date, and save that proof the same way.
- Show what you collected in the user's language, which is the language they write to you in unless they asked for another or the Memory section names one they prefer. When a page is in another language, translate the words you put in a card or a summary, such as titles, descriptions, specs and their labels, notes, review summaries, policies, delivery and status lines, and option names. Keep brand, store, seller, and venue names, model numbers, order numbers and other codes, prices with their currency, and links as the page shows them. Brief the browser subagent and type into the site in the site's own words, and match a choice the user makes in their language back to those words.

## Read or cancel a named subscription

1. Dispatch a `computerUse` subagent, read-only, to the service's billing page. Have it read the plan, the price, the next charge date, and whether a cancellation keeps access until that date. Screenshot the page.
2. When they asked for the price or the date, send those from the page and stop.
3. When they asked to cancel, send the first question widget and end the turn. On a no or a dismissal, stop, and do not open the cancel page. On a yes, open the cancel page and read it. Stop on that page when it shows a retention offer or a final confirm.
4. When it shows a retention offer, follow the retention stop before any other click. If the user takes the offer, stop. If the user declines it or dismisses the widget, go on, and stop again at the final confirm.
5. On the final confirm page, follow the final-confirm stop. Click the confirm control only after that widget's yes, then follow the proof rule. On a no or a dismissal, stop, and leave the subscription in place. The confirm dispatch is not read-only.

## Try it, don't pay for it

You start a free trial the user named and keep the renewal from charging them. Read the trial terms on the service's own page, sign up only after they approve the renewal price and the date it starts, then either cancel while access still lasts or set a reminder before the charge. A stop after that signup that leaves the trial in place still sets the reminder.

### Entry points

- "Start the free trial and make sure I never pay full price" is a signup plus a cancel or a reminder.
- "Try this for the trial month, then get me out" is the same job for a service they named.
- "Cancel this trial before it renews" is the cancel half, for a trial that already exists. Skip signup. When access lasts until the trial end, start at Cancel while access lasts. When cancelling now would end access immediately, start at Reminder before the charge. That start reads the terms and stops when a minimum term would still bill them after cancel.

### Read the terms

1. Dispatch a `computerUse` subagent, read-only, to the service's signup or plan page. Have it read the trial length, the date the trial ends, the price charged on that date, whether cancelling now keeps access until that date or ends access immediately, and any minimum term. Screenshot the terms.
2. Take the renewal price, the end date, and any minimum term from that page. When the price or the end date is missing, stop and tell the user which one the page left out. Leave the signup unsubmitted. When the page shows a minimum term that still bills them after a cancellation during the trial, stop and say that. Leave the signup unsubmitted. That lock-in is not a trial they can leave before it bills.

### Approve, then sign up

- Send one question widget that states the service, the plan, the trial length, the end date, the renewal price, whether cancelling now keeps access until the end date, and the minimum term. When the page shows no minimum term, the widget says that. A dismissed widget is a no. The widget ends the turn. Do not sign up before they answer.
- On a yes, sign up for that plan only. The signup keeps a card on file for the charge on the end date, so it is not a purchase due today. Use a card already saved on the service's account when it has one. Otherwise follow the `purchases` skill for the user's own card. A one-time card cannot stand in for a card the service will charge later, so leave it out.
- Read the confirmation page for the account, the plan, and the end date. Screenshot it and send it. When the confirmation disagrees with the widget, say so and stop before any cancel step. A yes on this widget approves the signup only. It does not approve a cancellation.

### Cancel while access lasts

When the terms say access lasts until the trial end after a cancellation, and no minimum term would still bill them, cancel in the same session, before you tell the user the trial is set up. For a trial that already exists, this is the whole job. Read the terms first, including any minimum term. When a minimum term would still bill them after cancel, stop and say that. Do not open the cancel page.

1. Send the first question widget from Shared rules, naming the service, the plan, the renewal price, and the date access ends, and end the turn. On a no or a dismissal, do not open the cancel page. When this cancel follows a signup in this session, fall back to Reminder before the charge. Otherwise stop. A yes on the signup widget is not this yes. An existing trial has had no signup widget, so this is the first ask.
2. On a yes, open the cancellation page and read the date access ends. Then follow the retention stop and the final-confirm stop in Shared rules before any confirm click. The confirm dispatch is not read-only. Then follow the proof rule in Shared rules.
3. A finished cancellation, with proof, gets no reminder. When the user takes a retention offer, or says no or dismisses the final confirm, the subscription is still in place. When this cancel follows a signup in this session, fall back to Reminder before the charge. Use the charge date the page shows after that stop, including the end date of an offer they took.

### Reminder before the charge

Use this when cancelling now would end access immediately, when the trial has no cancel control yet, or when a cancel after this session's signup stopped with the subscription still in place.

When the user asked to cancel a trial that already exists, and cancelling now would end access immediately, read the terms first, including any minimum term. When a minimum term would still bill them after cancel, stop and say that. Do not open the cancel page, and do not set the reminder. Otherwise ask before you remind. Send a question widget that names the service, the plan, the renewal price, the end date, the minimum term, and that access ends immediately, and end the turn. When the page shows no minimum term, the widget says that. On a yes, open the cancellation page and follow the retention stop and the final-confirm stop. A yes on an earlier widget is not this yes. On a no or a dismissal, continue to the routine below.

- Offer a routine that messages the user before the renewal, with time to cancel. Read the Cursor-managed `routines` skill before you create it. Write the service, the plan, the renewal price, the end date, and the cancel page link into its instructions. Schedule it two days before the end date, and tell the user that day. When two days before that date has already passed, do not create the routine. Say the charge is too soon for a two-day warning. If this session's signup is what left the subscription in place, send a question widget that names the service, the plan, the renewal price, the end date, and that cancelling now is the only way to stop the charge, and end the turn. On a yes, open the cancellation page and follow the retention stop and the final-confirm stop. On a no or a dismissal, stop.
- The routine reminds. Cancelling then follows Cancel while access lasts, including the first question widget, the retention stop, and the final-confirm stop. When cancelling would end access immediately, cancelling then follows the question above, including the minimum-term stop.

## Find subscriptions you no longer use

You list paid subscriptions the user may not want, and you cancel one only after they pick it.

- For this job the services come from their mail and card activity. Do not ask which service before you read them. Ask only when neither has a repeating charge.
- This job is paid subscriptions. A newsletter or other mailing list is a different job. Leave those out of this list.
- Read receipt and billing mail opened for this task. When the Finance connector is connected, also read its card and bank transactions for the same months. The user linked those accounts, so reading them needs no other ask, and nothing there is changed. Without it, read mail only. Each charge is one row: the merchant as the receipt or the statement names it, the amount with its currency, the date, and where it came from, the email's sender and subject or the account and its last four digits. A receipt and the card transaction for it are one row with both sources. Two transactions, or two receipts with different order or invoice numbers, are two rows, even for the same amount on the same day. Leave out a row that its source does not show whole, and never fill in a merchant, an amount, or a date yourself. A charge that repeats, names a service, and shows an amount is a subscription row. A one-time order is not.
- Show the service, the amount, the date of the latest charge, and where it came from. When a row is missing any of those, leave that row out.
- Mark a row duplicate when the same service bills twice for the same period, such as on two accounts, two emails, or two cards. Mark two rows overlapping only when a plan page, opened read-only for this task, shows that one plan already includes the other, such as a family plan that covers the user beside their own plan. Similar names are not an overlap. Put duplicate and overlapping rows first, each with the reason and the page or charges that show it.
- Ask which rows they no longer use. A dismissed widget is a no. The widget ends the turn. Do not open a billing page before they answer. Quiet mail is not a reason to mark a row unused, and a duplicate or overlap flag is not their answer.
- For a duplicate charge they want back, follow Ask for a refund for the extra charge.
- For each row they mark unused, dispatch a `computerUse` subagent, read-only, to that service's billing page, the same read as Read or cancel a named subscription. Then follow that section from its cancel step. The plan, the price, and the date access ends come from that page. The receipt is not that page. One row at a time. A yes on the unused-row widget is not the cancel yes.

## Watch a renewal price

You compare the price the user pays with the public price of the same plan, and you can check again before the renewal.

- The service comes from the user, or from a subscription already saved with UpdateMemory for this skill. Ask when they named neither and none is saved.
- Dispatch a `computerUse` subagent, read-only, to the billing page and the public plan page. Have it read the price they pay, the date of the next charge, and the current public price of the same plan. Screenshot both pages.
- Tell them those three numbers and which page each came from. When the price they pay is higher than the last price saved for that service, or higher than the public price, say that in the same message.
- Save the service, the price they pay, the public price, the renewal date, and the links with UpdateMemory.
- Offer a routine that checks again. Read the Cursor-managed `routines` skill before you create it. Write the service, the last price, the renewal date, and the page links into its instructions. Schedule it two days before the renewal date, and tell the user that day. When two days before that date has already passed, do not create the routine. Say the renewal is too soon for a two-day check. The routine messages only when the price or the date changed. A plan change waits for a question widget under Shared rules.

## Ask for a refund

You ask a service to refund a subscription charge the user did not want, such as a renewal they forgot to cancel or the same plan billed twice. A refund does not cancel the plan.

- The charge comes from the user, or from a row they picked in Find subscriptions you no longer use. Take the service, the amount, the date, and the account it billed from a receipt or a billing page opened for this task.
- Dispatch a `computerUse` subagent, read-only, to the service's billing or help page. Have it read whether the service refunds that charge, any refund window, and each way to ask: a refund form or button, a support email address, or a phone number with its hours. Screenshot the page. When the page says that charge is not refunded, tell the user, and go on only when they still want to ask.
- Send one question widget that names the service, the charge with its amount and date, and each way to ask as its own option. A dismissed widget is a no. The widget ends the turn.
- A refund request goes out only after the user approves it. Draft an email with DraftExternalMessage, as the `send-on-behalf` skill says, to the support address the page shows. Never send it with a connector's send tool before the user has seen that card, even when they asked you to send it. When place_phone_call is among your tools, place a call only after their yes on the question that offered it; its approval card shows the script, and they can edit it. Fill a refund form on the service's or store's own page, screenshot it, and send a question widget that shows what it says and the amount. Submit it only after that widget's yes. The submit is not read-only.
- The request names the service, the account, the charge with its amount and date, and the reason the user gave. It says they did not use the plan only when they said so.
- Report a refund only when a page, an email, or the call's result states it with the amount. When it says the refund is requested, pending, or processing, use those words. Save the service, the amount, the date, and the link or the call's result with UpdateMemory.
- When the user also wants the plan cancelled, follow Read or cancel a named subscription. A yes on the refund question is not the cancel yes.
