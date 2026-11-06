---
name: usage-review
description: >-
  When the user asks about their usage, limits, plan, or billing, or wants their
  bots and routines to use less, read this before you answer or change anything.
---
# Usage review

If you don't have CheckSubscriptionUsage, you can't see their numbers here, so answer what this skill covers without them and point them to Settings > Usage & Billing.
Answer the user's questions about usage, plans, and billing from real numbers, and suggest ways to use less.
If CheckSubscriptionUsage says usage wasn't checked, do what it says and stop.
Never state a usage figure, a price, or a reset time that the report or the user didn't give you.

## Look first
1. Call CheckSubscriptionUsage once. With no arguments it gives the plan, the usage used, on-demand, and when it resets, with the on-demand monthly limit and this billing cycle's on-demand amount when it shows them, which answers how much is used or left, when it resets, and whether they can keep working. For what used it or how to use less, group by bot and routine instead with `breakdown: ["bot", "routine"]`; it starts with the same lines. If the shares show little or no usage because the week just reset, call it again with `period: "last_7_days"`. Without grouping, low usage after a reset is the answer.
2. For why they hit their limit, also group by day with `breakdown: ["day"]` to find the day usage jumped.
3. For an earlier week, such as "how much did I use last week?", group by day with `breakdown: ["day"]` and `period: "last_30_days"`, or `"last_7_days"` for the last few days. For how much an Enterprise team's usage came to this billing cycle, pass `period: "this_billing_cycle"`. period is only this_week, last_7_days, last_30_days, or this_billing_cycle, so you can't pick an exact earlier week or billing cycle, or go back further than the last 30 days or the current billing cycle.
4. That is usually enough. Every read uses usage too. Read another bot's chat with ReadTranscript only when you can't judge its value from its name, at most two, with a small limit.

## Report
Start with how much of their usage is used and when it resets, as the report says it, then the on-demand monthly limit and amount when it shows them. When the report says usage counts toward their team's Enterprise usage, say that instead, and send them to their team's admin for what's left.
When you grouped usage, then name the two or three bots or routines with the largest shares.
Call things by the report's names, such as "weekly usage limit", "included usage", and "shares by bot and routine". Never name a tool, an option, or this skill, and never say an id, a folder, tokens, or events.
Use only the report's numbers, with dollars only when it shows them. The report can't predict, so never work out a saving or a future run count yourself. Describe a proposed schedule in words, such as "weekdays at 9 and 2".
Every share is of all Grok Bot usage, nested rows included, so a row's nested rows add up to its share. When the report says some usage isn't tied to one bot, the top-level rows add up to less than 100%; say so if they ask.
The percent used, on-demand, and reset are always current, whatever the period. Give an earlier week as its share of the window by adding up its days, such as "last week was about 40% of your Grok Bot usage over the last 30 days". Never call that a percent of last week's limit, and say that how much of the limit is used is shown only for the current week.
The report doesn't split usage by model, so you can't tell what a model used. Never blame or suggest one; answer from the bot, routine, and day shares.
When the report says the weekly limit is shared with other apps, say so, give Grok Bot's part as the report says it, and never guess what the other apps used. Otherwise included usage and the shares cover Grok Bot only. The on-demand amount, when the report shows it, covers Grok Bot and Cursor together, so say so. If they ask about Cursor or Cloud Agents usage on its own, say you can't see it here and point them to https://cursor.com/dashboard?tab=usage.

## Plans, billing, and estimates
For plans, prices, upgrades, cancelling, or on-demand, link the matching Settings row when your links list its anchor: usage, plan, cancel-trial, on-demand, or billing. The app is translated, so label the link in the user's language. Never quote a price from memory.
For a refund or a charge they think is wrong, tell them to email support@x.ai.
You can't predict what a task will use. Say what makes it larger, such as how many pages or sources it reads, how many steps it takes, and how often it repeats. Then offer to run it once and check the report afterwards.
For a warning before the limit, say the app already warns them as they near it, unless the report shows their plan is through their team, where it doesn't. Don't set up a routine to watch usage, since each check uses usage too. Offer one reminder instead, such as at the reset time from the report or before a renewal date they give.

## Suggest, then wait
Offer at most three changes, largest first. For each, say what changes, why, and what they'd give up. Ask which to make, then end your turn.
Make a change only after the user says yes to it, as the UpdateRoutine description also says. If they already named a specific change, make it.

## Changes to suggest
- Run a routine less often if it runs many times a week and its results don't need to be that fresh. They give up fresher results.
- Run a routine only on weekdays in waking hours if it runs at night or on weekends with nobody there to act on it. They give up off-hours answers, which wait until morning.
- Switch a routine from a schedule to a trigger if it checks for a Slack message, a GitHub or Origin change, an email, or another trigger the `routines` skill lists. They give up nothing in results, but it takes a short setup and that app must be connected.
- Merge two routines into one if they gather the same thing, often on different bots. They give up a separate routine for each job.
- Stop the work running now with StopSubagent all: true, and pause the routines that keep starting it if they want usage to stop right away. They give up the results of all of that work, including work you handed to other bots and cloud agents, since stopped work doesn't pick up again.
- Have them leave on-demand off, or set an on-demand spend limit themselves if the report shows on-demand and they want usage to stop at a set point, since nothing stops at a percentage of included usage. They give up any work past that point until the reset.
Never suggest changing Grok Bot's built-in check-ins.

## Make the change
Change your own routines with UpdateRoutine. Read one with action "get" first and edit its saved prompt rather than rewriting it. Pause rather than delete, so they can resume it.
You can't change another bot's routines. Tell the user which one to change in that bot's routine list, or, if they'd like, ask that bot with SendToAgent, which wakes it and uses usage too.
When the report shows on-demand, it's the user's to change, from the on-demand Settings row or the link in the report.
After a change, say what changed and how to undo it.
