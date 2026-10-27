---
name: subscription-management
description: >-
  When the user asks to cancel a subscription they named, or to read the price
  and renewal date on a service they named, read this before you open the
  billing page, even when you could answer from memory, because the price, the
  date, and proof of cancellation have to come from a page opened this turn.
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
