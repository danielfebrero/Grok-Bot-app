---
name: box-desktop
description: >-
  When a task needs your own desktop or browser, such as a website with no
  connector, a GUI app, or a sign-in only the user can complete, read before you
  dispatch the first browser or desktop subagent.
---
# The box desktop

You hold the read-only Screenshot tool to see its current screen, confirm where a flow landed, or check on a running subagent. You cannot click, move, type, press keys, scroll, or wait on the desktop yourself. Delegate every browser and desktop interaction to a subagent. Do not bypass this boundary with Shell-driven GUI automation such as xdotool, or by driving the box browser from Shell: no CDP attach, no Playwright, Puppeteer, or `websocket-client`, no `/json/new`, no cookie-DB scraping, and no page JS eval over DevTools.

- Delegate the outcome, constraints, required values, and success criteria, not browser or desktop steps. Prescribe a modality only when that modality is itself part of the desired result; otherwise let the child choose.

- When you know the destination URL, whether one the user pasted or one you can construct (a site's search/filter URL like `https://www.amazon.com/s?k=bread+flour`), put that exact URL in the task, as specific as the site's query params allow, so the subagent opens it directly instead of clicking through the site to rebuild it.

- Before dispatching to a website, check your skill catalog for that site's `site-playbooks-<site>` skill. If its description covers the job, Read it first and paste the matching section's Dispatch snippet into the task. That snippet's deep link and stop rules are the one case where you hand the subagent steps rather than only the outcome. Skip the Read when the job is not what the description names, or when that snippet is already in your instructions or earlier in this conversation.

- For bulk or structured data, don't type it in by hand: generate the file with Shell (e.g. a CSV), inspect it with Read when useful, then have the subagent import or upload it, far faster and more reliable than entering values one by one.

- When it returns, read its report before acting. If it stopped short or hit a step only the user can do, that's your cue to follow up or hand off the box.

- Some browser and desktop tasks take more than one or two minutes, for example a sign-up, a verification code, or a long form. Before you send one, tell the user in one line that it will take a few minutes. When its report comes back, tell the user what the report shows for that same task. Do not say a step happened, such as a code that was sent or a form that was submitted, until the report shows it.

- Sign in with the Cursor-managed `sign-in` skill, in its order: saved logins first, then any connected password manager, and only then the Secure Form or handing over the screen.

- When a page needs the USER to type (login, address, card, phone, OTP) and `RequestUserForm` is among your tools, Read and follow the `in-chat-forms` skill first. Have the child open the page and report fresh snapshot targets, then send the form; embedded and cross-origin card fields count. Do not jump to `request_box_help` just because a site needs a password.

- For steps that need the user, use `request_box_help` when `RequestUserForm` is not offered, a form cannot express the step (non-text/native steps), or the host refused the form or reported a failed fill. A structural preflight refusal is final, so don't re-issue the same form. Don't pre-ask "hand you the box?" The handoff is the ask.

- When a site in the task stops for a payment method, a paid plan, credits, a top-up, or a domain payment before it goes on, even inside a bigger job such as putting an app online, Read the Cursor-managed `purchases` skill first and follow it. Hand over the screen for that step only when that skill says to.

## Local egress

When the user turns on "Route traffic through this computer" (Settings → Computer in the Grok Bot app), your desktop browser's traffic goes out through their computer, so sites see their home or office network and IP instead of the cloud's. This helps with sites that block or challenge cloud traffic, region-specific content, and services that expect the user's usual location.

- It only applies to your desktop browser. Shell commands, WebFetch, and MCP connectors always go out from the cloud, so an IP check from Shell shows the cloud's address even when the setting is on.

- It works while the Grok Bot app is open on one of their computers. If they use several, it goes through one at a time, and Settings → Computer on each one shows whether that computer is the one routing.

- If a site keeps blocking you or asking for verification, you can suggest they turn the setting on.
