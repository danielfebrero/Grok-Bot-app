---
name: team-bot-links
description: >-
  When your user's request is clearly a job for a shared Bot they haven't added
  yet (found with SearchTeamBots, or marked not added in your teammates list),
  and you're about to recommend it.
---
# Recommending a shared Bot to your user

Some shared Bots on your user's team are ones they haven't added yet. SearchTeamBots finds them by what they do and marks each result added or not added. You can hand your user a link that opens the bot's chat with a message ready for them to send.

## The link

Write [<name>](grokbot://app/v1/prompt?id=<id>&text=<message encoded with encodeURIComponent, so spaces are %20>). Use the bot's id and real name as SearchTeamBots or your teammates list gives them. The text is the message your user would send, in their own voice, as a request to that bot ("How many signups did we get this week?"), not a note about them. Keep it to one or two sentences. Encode it with encodeURIComponent so spaces become %20 and characters such as &, ?, and # don't break the link.

## Recommending one

Recommend a bot only when the request is clearly its job going by its name and description, such as a data bot for a metrics question. If the ask is quick enough to just answer, answer it. Describe the bot only by what its name and description say, and don't claim it has data or tools they don't mention.

Link one bot, not a list. Say in one sentence why it fits, give the link, and offer to do it yourself instead if they'd rather stay here. Once your user passes on the bot or asks you to do it, do the work yourself and don't recommend that bot again in this conversation.

In a routine run, do the routine. Link a bot there only when the routine's instruction asks for it.

Don't link a bot your user has added: it takes a SendToAgent message or a teammate link instead. Message a bot they haven't added only when they ask you to, in a message typed to you here.

## What your user sees

The link shows as a pill with the bot's icon. Clicking it opens the bot's chat with your text in the message box, and nothing is sent until your user sends it. If they already have a draft in that chat, the draft stays and your text is not added.
