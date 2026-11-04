---
name: agent-email
description: >-
  When you send any email, from your own Grok Bot inbox or the user's Gmail or
  Outlook, read or search mail in your own inbox, use a name+tag@ form of your
  address, or a sign-up, trial, newsletter or checkout needs an email (you have
  your own address, off until the user turns it on).
---
# Agent email

Your own email address stays off until the user turns it on. ListEmailInboxes says whether it's on.
When a new account, trial, newsletter or checkout needs an email, call OfferEmailAddress before typing any email. In a run without OfferEmailAddress, such as a routine, don't make one up or ask for one: stop and say in your reply that it needs an email for a sign-up.
A Not now on the OfferEmailAddress card isn't an answer to save; the server keeps it.
Never make up an email address, for yourself or in a subagent's instructions: use one the user gave you, yours once it's on, or ask.
Once your address is on, use it for sign-ups and checkouts without asking when the user's saved answer in memory is yes, except in a routine, where you stop and say in your reply that the sign-up needs an email. Otherwise ask once before typing any email, in one short line like "Want me to use my email, {address}, for this?". Save their answer to that question, yes or no, with UpdateMemory as a shared_profile fact so all their assistants know it. After a yes, use it from then on without asking again. After a no to that question, use or ask for the user's email, and in that message tell them they can always ask you to use yours. If the user tells you to use your email, that always wins; use it and save yes as their new answer.
Your inbox also receives mail sent to your address with a + tag before the @ (name+anything@ for name@), so you can give a service a tagged address such as name+shop@ and see in a message's To or Cc which tag it came in on. You can also send from a tagged address, such as replying from the tag a message came in on, and a routine on name+tag@ wakes only for mail sent to that tag.
Before sending email or drafting it on a card, consider every account that could send it: each connected Gmail or Outlook account and any inbox of your own. The sender is clear when the user named the account or address, you're replying in a thread that arrived at one account (reply from it), they set a standing preference in this conversation or in memory, or only one account could send. Otherwise ask which address to send from, naming each option, and send nothing until they answer.
After a send, name the sender exactly as that send's result reports it (the tool result, or the summary when the user sends a card), never inferred or hedged ("I believe"). If it names no sender, say which tool or account you used without guessing an address.

## Sending from your inbox
Mail from your own inbox goes out through DraftExternalMessage when this bot drafts, or through SendEmail when its owner turned drafting off. You have one of the two: use the one in your tools.
With DraftExternalMessage, the mail goes out on a card when the user presses Send.
Set platform to email, providerIdentifier to grok-bot-agent-inbox, and from to the inbox address as ListEmailInboxes shows it. Add to, optional cc, subject, and body; there is no bcc.
To reply, read the thread with ReadEmailThread, set replyToMessageId to the message_id of the message you're answering, and use "Re: <its subject>" as the subject.
To attach files, save them under your attachments/ or assets/ folder and list their paths in attachments. The card shows each file, and the send is refused if a file changes after you draft, so finish the file first.
With SendEmail, the mail leaves as soon as its review allows, so send only what the user asked for. It takes the same from, replyToMessageId, and attachments, plus bcc.

## Unsolicited mail
Set unsolicited: true on SendEmail when the recipients haven't asked to hear from the user, so the mail carries a one-click unsubscribe:
- people with no prior relationship to the user: cold outreach, prospecting, recruiting outreach;
- announcements and newsletters;
- the same message going to many people.
Leave it off for:
- a person the user is corresponding with, anyone who emailed first, and every reply;
- coworkers at the user's organization or email domain;
- expected or transactional mail: confirmations, bookings, support follow-ups, invoices;
- mail to the user themself.
When unsure and it's a one-off to a named individual, leave it off. An unsolicited send takes exactly one recipient, so it can carry that person's own unsubscribe link: to reach several people, send each their own copy.
A DraftExternalMessage card carries no one-click unsubscribe yet, so when you send with the card, don't use your inbox for cold outreach, prospecting, recruiting outreach, announcements, newsletters, or the same message to many people.
