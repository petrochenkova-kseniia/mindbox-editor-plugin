#!/usr/bin/env python3
"""
Trello helper — boards, lists, and cards management.

Usage:
  python3 trello_helper.py boards                          # list all boards
  python3 trello_helper.py lists BOARD_ID                  # list all lists on a board
  python3 trello_helper.py cards LIST_ID                   # list cards in a list
  python3 trello_helper.py card CARD_ID                    # show card details
  python3 trello_helper.py create LIST_ID "Card name"      # create a card
  python3 trello_helper.py move CARD_ID LIST_ID            # move card to another list
  python3 trello_helper.py comment CARD_ID "Text"          # add comment to a card
  python3 trello_helper.py search "query"                  # search cards

  python3 trello_helper.py copy CARD_ID LIST_ID "Name"  # copy card to list
  python3 trello_helper.py update-desc CARD_ID           # update desc (from stdin)
  python3 trello_helper.py history CARD_ID               # full history: list movements + comments
"""

import json
import os
import sys
import urllib.request
import urllib.parse

CREDS_PATH = os.path.expanduser("~/Claude/integrations/credentials/trello.json")
API_BASE = "https://api.trello.com/1"


def load_creds():
    if not os.path.exists(CREDS_PATH):
        print(f"ERROR: {CREDS_PATH} not found.")
        sys.exit(1)
    with open(CREDS_PATH) as f:
        return json.load(f)


def api_request(method, path, params=None, data=None):
    creds = load_creds()
    if params is None:
        params = {}
    params["key"] = creds["api_key"]
    params["token"] = creds["token"]

    url = f"{API_BASE}{path}"
    if method == "GET":
        url += "?" + urllib.parse.urlencode(params)
        req = urllib.request.Request(url)
    else:
        req = urllib.request.Request(
            url + "?" + urllib.parse.urlencode(params),
            data=json.dumps(data).encode() if data else None,
            method=method,
        )
        req.add_header("Content-Type", "application/json")

    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        print(f"ERROR {e.code}: {body}")
        sys.exit(1)


def cmd_boards():
    boards = api_request("GET", "/members/me/boards", {"fields": "name,url,closed"})
    for b in boards:
        if not b.get("closed"):
            print(f"ID: {b['id']}")
            print(f"  {b['name']}")
            print(f"  {b['url']}")
            print()


def cmd_lists(board_id):
    lists = api_request("GET", f"/boards/{board_id}/lists", {"fields": "name,pos"})
    for lst in lists:
        print(f"ID: {lst['id']}  —  {lst['name']}")


def cmd_cards(list_id):
    cards = api_request("GET", f"/lists/{list_id}/cards", {
        "fields": "name,desc,due,labels,url,idMembers",
    })
    if not cards:
        print("No cards in this list.")
        return
    for c in cards:
        labels = ", ".join(l["name"] or l["color"] for l in c.get("labels", []))
        due = c.get("due", "")
        print(f"ID: {c['id']}")
        print(f"  {c['name']}")
        if labels:
            print(f"  Labels: {labels}")
        if due:
            print(f"  Due: {due}")
        print()


def cmd_card(card_id):
    c = api_request("GET", f"/cards/{card_id}", {
        "fields": "name,desc,due,labels,url,idList,idBoard",
        "actions": "commentCard",
        "actions_limit": "10",
    })
    print(f"Name: {c['name']}")
    print(f"URL:  {c['url']}")
    if c.get("desc"):
        print(f"Desc: {c['desc'][:200]}")
    if c.get("due"):
        print(f"Due:  {c['due']}")
    labels = ", ".join(l["name"] or l["color"] for l in c.get("labels", []))
    if labels:
        print(f"Labels: {labels}")
    print()
    actions = c.get("actions", [])
    if actions:
        print("--- Comments ---")
        for a in actions:
            author = a["memberCreator"]["fullName"]
            text = a["data"]["text"]
            date = a["date"][:10]
            print(f"  [{date}] {author}: {text}")


def cmd_create(list_id, name):
    card = api_request("POST", "/cards", {"idList": list_id, "name": name})
    print(f"Created: {card['url']}")


def cmd_move(card_id, list_id):
    api_request("PUT", f"/cards/{card_id}", {"idList": list_id})
    print(f"Card {card_id} moved to list {list_id}")


def cmd_comment(card_id, text):
    api_request("POST", f"/cards/{card_id}/actions/comments", {"text": text})
    print("Comment added.")


def cmd_copy(card_id, list_id, name):
    card = api_request("POST", "/cards", {
        "idList": list_id,
        "idCardSource": card_id,
        "name": name,
        "keepFromSource": "checklists",
    })
    print(f"Created: {card['id']}")
    print(f"URL: {card['url']}")


def cmd_update_desc(card_id):
    desc = sys.stdin.read()
    creds = load_creds()
    url = (
        f"{API_BASE}/cards/{card_id}"
        f"?key={creds['api_key']}&token={creds['token']}"
    )
    req = urllib.request.Request(
        url,
        data=json.dumps({"desc": desc}).encode(),
        method="PUT",
    )
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req) as resp:
            resp.read()
    except urllib.error.HTTPError as e:
        print(f"ERROR {e.code}: {e.read().decode()}")
        sys.exit(1)
    print(f"Description updated for card {card_id}")


def cmd_history(card_id):
    card = api_request("GET", f"/cards/{card_id}", {"fields": "name,url,idList,idBoard"})
    print(f"Card: {card['name']}")
    print(f"URL:  {card['url']}")
    print()

    actions = api_request(
        "GET",
        f"/cards/{card_id}/actions",
        {"filter": "createCard,updateCard:idList,commentCard,copyCard,convertToCardFromCheckItem", "limit": "1000"},
    )

    events = []
    for a in actions:
        t = a.get("type")
        date = a["date"]
        author = a.get("memberCreator", {}).get("fullName", "?")
        if t == "createCard":
            lst = a.get("data", {}).get("list", {}).get("name", "?")
            events.append((date, "create", f"Created in '{lst}'", author))
        elif t == "copyCard":
            lst = a.get("data", {}).get("list", {}).get("name", "?")
            events.append((date, "create", f"Copied into '{lst}'", author))
        elif t == "convertToCardFromCheckItem":
            lst = a.get("data", {}).get("list", {}).get("name", "?")
            events.append((date, "create", f"Converted into '{lst}'", author))
        elif t == "updateCard":
            data = a.get("data", {})
            before = data.get("listBefore", {}).get("name", "?")
            after = data.get("listAfter", {}).get("name", "?")
            events.append((date, "move", f"'{before}' -> '{after}'", author))
        elif t == "commentCard":
            text = a.get("data", {}).get("text", "")
            events.append((date, "comment", text, author))

    events.sort(key=lambda x: x[0])

    print("=== Movements between lists ===")
    moves = [e for e in events if e[1] in ("create", "move")]
    if not moves:
        print("  (none)")
    for date, kind, text, author in moves:
        print(f"  [{date[:10]}] {text}  (by {author})")
    print()

    print("=== Comments ===")
    comments = [e for e in events if e[1] == "comment"]
    if not comments:
        print("  (none)")
    for date, kind, text, author in comments:
        print(f"  [{date[:10]}] {author}: {text}")


def cmd_search(query):
    results = api_request("GET", "/search", {
        "query": query,
        "modelTypes": "cards",
        "cards_limit": "20",
        "card_fields": "name,url,idList,due",
    })
    cards = results.get("cards", [])
    if not cards:
        print(f"No cards found for: {query}")
        return
    print(f"Found {len(cards)} card(s):\n")
    for c in cards:
        print(f"ID: {c['id']}")
        print(f"  {c['name']}")
        print(f"  {c['url']}")
        print()


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    command = sys.argv[1]
    if command == "boards":
        cmd_boards()
    elif command == "lists":
        if len(sys.argv) < 3:
            print("Usage: trello_helper.py lists BOARD_ID")
            sys.exit(1)
        cmd_lists(sys.argv[2])
    elif command == "cards":
        if len(sys.argv) < 3:
            print("Usage: trello_helper.py cards LIST_ID")
            sys.exit(1)
        cmd_cards(sys.argv[2])
    elif command == "card":
        if len(sys.argv) < 3:
            print("Usage: trello_helper.py card CARD_ID")
            sys.exit(1)
        cmd_card(sys.argv[2])
    elif command == "create":
        if len(sys.argv) < 4:
            print('Usage: trello_helper.py create LIST_ID "Card name"')
            sys.exit(1)
        cmd_create(sys.argv[2], sys.argv[3])
    elif command == "move":
        if len(sys.argv) < 4:
            print("Usage: trello_helper.py move CARD_ID LIST_ID")
            sys.exit(1)
        cmd_move(sys.argv[2], sys.argv[3])
    elif command == "comment":
        if len(sys.argv) < 4:
            print('Usage: trello_helper.py comment CARD_ID "Text"')
            sys.exit(1)
        cmd_comment(sys.argv[2], sys.argv[3])
    elif command == "copy":
        if len(sys.argv) < 5:
            print('Usage: trello_helper.py copy CARD_ID LIST_ID "Name"')
            sys.exit(1)
        cmd_copy(sys.argv[2], sys.argv[3], sys.argv[4])
    elif command == "update-desc":
        if len(sys.argv) < 3:
            print("Usage: echo 'desc' | trello_helper.py update-desc CARD_ID")
            sys.exit(1)
        cmd_update_desc(sys.argv[2])
    elif command == "search":
        if len(sys.argv) < 3:
            print('Usage: trello_helper.py search "query"')
            sys.exit(1)
        cmd_search(sys.argv[2])
    elif command == "history":
        if len(sys.argv) < 3:
            print("Usage: trello_helper.py history CARD_ID")
            sys.exit(1)
        cmd_history(sys.argv[2])
    else:
        print(f"Unknown command: {command}")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
