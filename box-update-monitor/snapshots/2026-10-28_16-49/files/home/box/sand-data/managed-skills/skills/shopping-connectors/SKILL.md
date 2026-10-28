---
name: shopping-connectors
description: >-
  When the user wants to shop for, find, compare, or order something, before you
  search for it, so you offer and use their shopping connectors.
---
# Shopping connectors

Two connectors can search and order for the user. Gopuff covers groceries, snacks, drinks, household items, and other everyday goods delivered in minutes. Shop searches products across many merchants, and keeps carts, checkout, and order history for the user's Shop account. A connector is connected when its plugin's MCP tools are among your tools this turn. Once connected, its tools, not the open web, are the place to look for what it sells.
- When the user wants to shop for something, check GetMcpServerStatus for each connector that could plausibly fill the request. For each one that is not installed, offer it once, in one short friendly line, before you search elsewhere: say what it can search and order, and ask if they want to connect it. Offer both in the same line when both fit. Do not offer one again in this conversation if they decline or ignore it.
- If they agree, call InstallPlugin with plugin_id 77083001 for Gopuff or 77565631 for Shop. If it is already installed but its server needsAuth, call AuthenticateMcpServer on that server instead. Never write a link to it yourself: in the app its connect card resumes you when they finish, and elsewhere the result says what to relay.
- When a connector is connected, search it for every shopping request it could plausibly fill, alongside your other ways to search for products rather than after them, and compare what each returns. Read that plugin's own skills and tool descriptions before its first action and follow them for addresses, carts, and orders.
- Name where each result came from, and say plainly when a connector has no match instead of leaving it out. Prices, availability, and delivery depend on the delivery address, so confirm the address before you quote them.
- Searching is free to do on your own, but changing a cart or placing an order is not: ask before you add anything to a cart, and place an order only after the user approves that exact order and its total.
- Some connectors take payment in their own tools: Gopuff quotes an order with get_quote and charges it with place_order. Pay those orders only with a way to pay that the connector's quote lists, such as a payment method saved on the user's account there or cash where offered. Never pay one with a card from RequestVirtualCard, Shop, X Money, or the Secure Form: those cards are typed into a web checkout page, and a connector order has none. A Stripe Link option on the quote needs a Link shared payment token, which you can't get, so don't pick it. When nothing on the quote can pay, say so in one line and give the user the connector's own link to add a payment method or finish in its app, from its tool results. Any other connector order follows the `purchases` skill.
- When the user declines a connector or it cannot fill the request, carry on with the other ways to shop without bringing that connector up again.
