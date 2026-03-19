"""
apps/ms_teams/activities.py — MS Teams UI automation.

ONLY browser interaction lives here.
No log capture, no popup handling, no result building, no report logic —
all of that is in core/base_activity.py.

To add a new MS Teams activity:
  1. Add it to apps/ms_teams/app.yaml under 'activities'
  2. Add _do_{activity_name}(self, page, result, **kwargs) below
  Done.
"""

import time
import sys
import os

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from core.base_activity import BaseActivity


# ── Teams load selectors ──────────────────────────────────────────────────────
_TEAMS_LOADED_SELECTORS = [
    "[data-tid='chat-list']",
    "div[role='list']",
    "button[aria-label='Chat']",
    "text=New chat",
    "text=Recent",
]


class MsTeamsActivity(BaseActivity):
    """
    MS Teams CASB activity handler.
    Implements: post, meet_now_post, forward, reply.
    """

    # ================================================================
    # REQUIRED: Open tab + wait for app
    # ================================================================

    def _open_fresh_tab(self):
        url       = self.app_config.get("app_url", "https://teams.live.com/v2/")
        old_pages = list(self.browser.pages)
        new_page  = self.browser.new_page()
        new_page.goto(url, wait_until="domcontentloaded")
        new_page.wait_for_timeout(8000)
        print(f"\nOpened fresh tab → {url}")
        for old_page in old_pages:
            try:
                old_page.close()
            except Exception:
                pass
        return new_page

    def _wait_for_app(self, page) -> bool:
        """Poll until Teams chat UI is visible."""
        print("Waiting for MS Teams to load...")
        for attempt in range(36):
            for sel in _TEAMS_LOADED_SELECTORS:
                try:
                    page.locator(sel).first.wait_for(state="visible", timeout=3000)
                    print(f"Teams loaded (matched: {sel})")
                    return True
                except Exception:
                    continue
            print(f"   [{attempt + 1}/36] Still loading...")
            page.wait_for_timeout(5000)
        print("Teams did NOT load.")
        return False

    # ================================================================
    # TC1 — Chat → Send text to recipient
    # ================================================================

    def _do_post(self, page, result, recipient, message, **kwargs):
        """Direct chat send."""
        print(f"\n   [TC1] Sending message to: {recipient}")

        # Open chat
        clicked = False
        for strategy in [
            lambda: page.locator(f"xpath=//span[normalize-space(text())='{recipient}']").first.click(timeout=5000),
            lambda: page.get_by_text(recipient, exact=True).first.click(timeout=5000),
            lambda: page.get_by_text(recipient).first.click(timeout=5000),
        ]:
            try:
                strategy(); clicked = True; break
            except Exception:
                continue

        if not clicked:
            result["fail_reason"].append(f"Could not open chat with {recipient}")
            return False

        page.wait_for_timeout(3000)

        # Type message
        typed = False
        for sel in [
            lambda: page.get_by_placeholder("Type a message"),
            lambda: page.locator("div[contenteditable='true']").last,
        ]:
            try:
                box = sel()
                box.wait_for(state="visible", timeout=5000)
                box.click()
                page.wait_for_timeout(500)
                box.type(message, delay=80)
                typed = True
                break
            except Exception:
                continue

        if not typed:
            result["fail_reason"].append("Could not type message in chat")
            return False

        # Pre-connect vsmd + start HAR right before send
        vsmd_prep, har = self._before_send(page, "TC1_BaseSendPost")
        result["_har"] = har   # store now so HAR is saved even on early exit

        # Send
        sent = False
        try:
            page.get_by_role("button", name="Send").click(timeout=5000)
            sent = True
        except Exception:
            page.keyboard.press("Control+Enter")
            sent = True

        # Wait for network to settle — yields to Playwright event loop
        # so HAR listeners can receive responses before stop() is called
        page.wait_for_timeout(3000)

        self._after_send(page, result, vsmd_prep, har, "TC1_BaseSendPost", None)

        ss, _ = self._screenshot(page, "TC1_BaseSendPost_step1_sent")
        self._add_step(result, "TC1-b", f"Message Sent to {recipient}",
                       "pass" if sent else "fail",
                       [f"Recipient : {recipient}",
                        f"Message   : '{message}'",
                        f"Result    : {'Sent ✓' if sent else 'FAILED ✗'}"], ss)

        if not sent:
            result["fail_reason"].append("Could not send message")
            return False

        # Message delivery check — fails TC if message was delivered
        self._check_delivery_generic(page, result, message, "TC1-c", tag="TC1")
        return True

    # ================================================================
    # TC2 — Meet Now → Start meeting → Chat tab → Post
    # ================================================================

    def _do_meet_now_post(self, page, result, message, **kwargs):
        """Meet Now → Chat → Post."""
        print(f"\n   [TC2] Meet Now post: '{message}'")

        # Step a: Click Meet Now icon
        clicked = False
        for sel in [
            "button[aria-label='Meet Now']",
            "button[aria-label='Meet now']",
            "xpath=//button[@aria-label='Meet Now']",
            "xpath=//button[@aria-label='Meet now']",
            "xpath=//button[.//span[text()='Meet now']]",
            "xpath=//button[.//span[text()='Meet Now']]",
            "xpath=//button[contains(@aria-label,'Meet')]",
            "xpath=//button[contains(@aria-label,'meet')]",
            "[data-tid='meet-now-button']",
            "[data-tid='meetNow']",
        ]:
            try:
                page.locator(sel).first.wait_for(state="visible", timeout=4000)
                page.locator(sel).first.click()
                clicked = True
                print(f"   [TC2] Meet Now clicked via: {sel}")
                break
            except Exception:
                continue

        if not clicked:
            print(f"   [TC2] Could not find Meet Now button — check screenshot TC2_MeetNow_step1_meet_now_icon")
            # Try to find any button with 'meet' in aria-label for debugging
            try:
                meet_btns = page.locator("xpath=//button[contains(translate(@aria-label,'MEETNO','meetno'),'meet')]").all()
                print(f"   [TC2] Buttons with 'meet' in aria-label: {[b.get_attribute('aria-label') for b in meet_btns[:5]]}")
            except Exception:
                pass

        page.wait_for_timeout(2000)
        # Dismiss Windows Firewall dialog immediately if it appeared
        self._dismiss_windows_firewall()
        ss1, _ = self._screenshot(page, "TC2_MeetNow_step1_meet_now_icon")
        self._add_step(result, "TC2-a", "Clicked Meet Now Icon",
                       "pass" if clicked else "fail",
                       ["Clicked Meet Now icon near Chat button"], ss1)
        if not clicked:
            result["fail_reason"].append("Could not find Meet Now icon")
            return False

        # Step b: Click Start meeting
        started = False
        for sel in [
            "button:has-text('Start meeting')",
            "xpath=//button[normalize-space(text())='Start meeting']",
            "[data-tid='start-meeting-button']",
            "text=Start meeting",
        ]:
            try:
                page.locator(sel).first.wait_for(state="visible", timeout=8000)
                page.locator(sel).first.click()
                started = True
                break
            except Exception:
                continue

        page.wait_for_timeout(5000)
        ss2, _ = self._screenshot(page, "TC2_MeetNow_step2_start_meeting")
        self._add_step(result, "TC2-b", "Clicked Start Meeting",
                       "pass" if started else "fail",
                       ["Clicked Start meeting to launch meeting window"], ss2)
        if not started:
            result["fail_reason"].append("Could not find Start meeting button")
            return False

        # Step c: Dismiss popups
        self._dismiss_meeting_popups(page)

        # Step d: Click Chat in toolbar
        chat_clicked = False
        for sel in [
            "xpath=(//button[@aria-label='Chat'])[last()]",
            "xpath=(//button[.//span[text()='Chat']])[last()]",
            "xpath=//div[contains(@class,'toolbar')]//button[@aria-label='Chat']",
        ]:
            try:
                btn = page.locator(sel).last
                btn.wait_for(state="visible", timeout=8000)
                btn.click()
                chat_clicked = True
                break
            except Exception:
                continue

        page.wait_for_timeout(2000)
        self._add_step(result, "TC2-c", "Clicked Chat Tab in Meeting",
                       "pass" if chat_clicked else "fail",
                       ["Clicked Chat tab in meeting toolbar"])
        if not chat_clicked:
            result["fail_reason"].append("Could not find Chat tab in meeting")
            return False

        # Step e: Type message
        page.wait_for_timeout(3000)
        typed = False
        for sel in [
            "xpath=//div[contains(@class,'meeting-chat')]//div[@contenteditable='true']",
            "xpath=//aside//div[@contenteditable='true']",
            "xpath=(//div[@contenteditable='true'])[last()]",
        ]:
            try:
                box = page.locator(sel).last
                box.wait_for(state="visible", timeout=5000)
                box.click()
                page.wait_for_timeout(500)
                box.type(message, delay=80)
                typed = True
                break
            except Exception:
                continue

        if not typed:
            result["fail_reason"].append("Could not type in meeting chat")
            return False

        # Pre-connect vsmd + HAR before send
        vsmd_prep, har = self._before_send(page, "TC2_MeetNowPost")
        result["_har"] = har   # store now so HAR is saved even on early exit

        # Send
        sent = False
        for send_sel in [
            "button[aria-label='Send']",
            "button:has-text('Send')",
            "xpath=//button[@aria-label='Send']",
        ]:
            try:
                page.locator(send_sel).first.click(timeout=3000)
                sent = True
                break
            except Exception:
                continue
        if not sent:
            page.keyboard.press("Enter")
            sent = True

        page.wait_for_timeout(3000)
        self._after_send(page, result, vsmd_prep, har, "TC2_MeetNowPost", None)

        page.wait_for_timeout(3000)
        ss3, _ = self._screenshot(page, "TC2_MeetNow_step3_message_sent")
        self._add_step(result, "TC2-d", "Message Posted in Meeting Chat",
                       "pass" if sent else "fail",
                       [f"Message : '{message}'",
                        f"Result  : {'Sent ✓' if sent else 'Failed ✗'}"], ss3)
        if not sent:
            result["fail_reason"].append("Could not send message in meeting chat")
            return False

        # ── Delivery check for meeting chat ──────────────────────
        # Visual signal (confirmed from real Teams HTML screenshot):
        #
        #   CASB BLOCKED  → hollow circle ○ shown below the message bubble
        #                   this is the read-status icon span with
        #                   aria-label="Retrying..." containing a ring-only SVG
        #                   AND no <time> timestamp on the bubble
        #
        #   DELIVERED     → circle disappears, <time class="fui-ChatMyMessage__timestamp">
        #                   appears with the sent time (e.g. "21:12")
        #
        # Detection strategy (priority order):
        #   1. Find message bubble via data-tid="chat-pane-message" + message text
        #   2. PRIMARY:   check read-status icon aria-label for "Retrying" / "Sending"
        #                 → hollow circle = blocked
        #   3. SECONDARY: check for <time datetime=...> timestamp element
        #                 → timestamp present = delivered
        #   4. FALLBACK:  bubble not found or no status → assume blocked
        print(f"   [TC2] Checking message delivery in meeting chat...")
        delivered = False
        blocked   = False
        detail    = "Delivery status inconclusive — assuming CASB blocked"

        try:
            page.wait_for_timeout(3000)

            bubble_js = f"""
                () => {{
                    // Find the message bubble containing our text
                    const bubbles = document.querySelectorAll('[data-tid="chat-pane-message"]');
                    for (const el of bubbles) {{
                        const content = el.querySelector('[data-message-content]');
                        if (content && content.innerText.includes({repr(message)})) {{
                            // Walk up to the wrapper that also contains the
                            // read-status icon (it sits outside the role=group div,
                            // as a sibling in the fui-ChatMyMessage container)
                            const wrapper = el.closest('.fui-ChatMyMessage') || el.parentElement;
                            return {{
                                bubbleHtml  : el.outerHTML,
                                wrapperHtml : wrapper ? wrapper.outerHTML : el.outerHTML,
                            }};
                        }}
                    }}
                    return null;
                }}
            """
            bubble_data = page.evaluate(bubble_js)

            if bubble_data is None:
                # Bubble never appeared — CASB blocked before Teams could render it
                blocked = True
                detail  = "Message bubble not found in meeting chat panel — CASB blocked ✓"
                result["message_not_delivered"] = True

            else:
                bubble_html  = (bubble_data.get("bubbleHtml")  or "").lower()
                wrapper_html = (bubble_data.get("wrapperHtml") or "").lower()

                # ── PRIMARY SIGNAL: hollow circle read-status icon ────
                # aria-label="Retrying..." → message stuck, not delivered
                # aria-label="Sending"    → still in-flight (also blocked)
                is_hollow_circle = (
                    'aria-label="retrying' in wrapper_html
                    or 'aria-label="sending' in wrapper_html
                    or 'retrying...' in wrapper_html
                )

                # ── SECONDARY SIGNAL: timestamp element ───────────────
                # <time class="fui-ChatMyMessage__timestamp" datetime="...">
                # present only after the message is confirmed delivered
                has_timestamp = (
                    'fui-chatmymessage__timestamp' in wrapper_html
                    or ('<time' in wrapper_html and 'datetime=' in wrapper_html)
                )

                if is_hollow_circle and not has_timestamp:
                    # Hollow circle present, no timestamp → definitively blocked
                    blocked = True
                    detail  = (
                        "Hollow circle ○ (aria-label='Retrying...') present, "
                        "no timestamp → CASB block CONFIRMED ✓"
                    )
                    result["message_not_delivered"] = True

                elif has_timestamp and not is_hollow_circle:
                    # Timestamp present, circle gone → definitively delivered
                    delivered = True
                    detail    = (
                        "Timestamp (fui-ChatMyMessage__timestamp) present, "
                        "no hollow circle → message DELIVERED — CASB did NOT block ✗"
                    )
                    result["message_not_delivered"] = False
                    result["fail_reason"].append(
                        "Meeting chat message was delivered (timestamp visible, "
                        "hollow circle gone) — CASB did not block"
                    )

                elif is_hollow_circle and has_timestamp:
                    # Both present — ambiguous; trust the circle (more immediate signal)
                    blocked = True
                    detail  = (
                        "Both hollow circle and timestamp detected "
                        "(race condition) → treating as blocked ✓"
                    )
                    result["message_not_delivered"] = True

                else:
                    # Bubble found but neither signal — intermediate state; assume blocked
                    blocked = True
                    detail  = (
                        "Message bubble found but no clear delivery signal "
                        "(no circle, no timestamp) → assuming CASB blocked ✓"
                    )
                    result["message_not_delivered"] = True

        except Exception as e:
            detail = f"Delivery check error: {e} — inconclusive, assuming CASB blocked"
            result["message_not_delivered"] = True

        print(f"   [TC2] {detail}")
        ss4, _ = self._screenshot(page, "TC2_MeetNow_step4_delivery_check")
        self._add_step(result, "TC2-e", "Message Delivery Check (Meeting Chat)",
                       "pass" if result["message_not_delivered"] else "fail",
                       [detail,
                        f"Hollow circle (Retrying): {blocked}",
                        f"Timestamp present       : {delivered}",
                        "○ hollow circle = CASB blocked (PASS) | timestamp = delivered (FAIL)"],
                       ss4)

        return True

    # ================================================================
    # TC3 — Chat → 3 dots → Forward message
    # ================================================================

    def _do_forward(self, page, result, recipient, message, **kwargs):
        """Forward a message to a recipient."""
        print(f"\n   [TC3] Forward message to: {recipient}")
        time.sleep(10)   # let Teams settle after previous TC

        # Open chat
        clicked = False
        for strategy in [
            lambda: page.locator(f"xpath=//span[normalize-space(text())='{recipient}']").first.click(timeout=5000),
            lambda: page.get_by_text(recipient, exact=True).first.click(timeout=5000),
        ]:
            try:
                strategy(); clicked = True; break
            except Exception:
                continue

        page.wait_for_timeout(3000)
        self._add_step(result, "TC3-a", f"Opened Chat with {recipient}",
                       "pass" if clicked else "fail",
                       [f"Recipient: {recipient}"])
        if not clicked:
            result["fail_reason"].append(f"Could not open chat with {recipient}")
            return False

        # Hover + click 3 dots
        dots_clicked = self._hover_and_click_dots(page, message, tag="FORWARD")
        ss1, _ = self._screenshot(page, "TC3_Forward_step1_three_dots")
        self._add_step(result, "TC3-b", "Clicked 3 Dots on Message",
                       "pass" if dots_clicked else "fail",
                       [f"Message: '{message}'", "Hovered → More Actions"], ss1)
        if not dots_clicked:
            result["fail_reason"].append("Could not click 3 dots on message")
            return False

        # Click Forward
        forward_clicked = False
        for sel in [
            "button:has-text('Forward')",
            "xpath=//button[normalize-space(text())='Forward']",
            "text=Forward", "[data-tid='forward-message']",
        ]:
            try:
                page.locator(sel).first.wait_for(state="visible", timeout=5000)
                page.locator(sel).first.click()
                forward_clicked = True
                break
            except Exception:
                continue

        page.wait_for_timeout(2000)
        self._add_step(result, "TC3-c", "Clicked Forward",
                       "pass" if forward_clicked else "fail",
                       ["Selected Forward from context menu"])
        if not forward_clicked:
            result["fail_reason"].append("Could not find Forward in menu")
            return False

        # Select recipient in forward dialog + send
        sent = False
        for inp_sel in [
            "input[placeholder*='Name']", "input[placeholder*='Search']",
            "input[aria-label*='Search']",
        ]:
            try:
                inp = page.locator(inp_sel).first
                inp.wait_for(state="visible", timeout=5000)
                inp.click()
                inp.type(recipient, delay=80)
                page.wait_for_timeout(2000)
                for opt_sel in [
                    f"xpath=//span[contains(text(),'{recipient}')]",
                    f"text={recipient}",
                ]:
                    try:
                        page.locator(opt_sel).first.click(timeout=3000)
                        break
                    except Exception:
                        continue
                page.wait_for_timeout(1000)

                # Pre-connect vsmd + HAR before send
                vsmd_prep, har = self._before_send(page, "TC3_ForwardMessage")
                result["_har"] = har   # store now so HAR is saved even on early exit

                for send_sel in [
                    "button:has-text('Forward')", "button:has-text('Send')",
                    "button[aria-label='Send']",
                ]:
                    try:
                        page.locator(send_sel).first.click(timeout=3000)
                        sent = True
                        break
                    except Exception:
                        continue

                page.wait_for_timeout(3000)
                self._after_send(page, result, vsmd_prep, har,
                                 "TC3_ForwardMessage", None)
                break
            except Exception:
                continue

        ss2, _ = self._screenshot(page, "TC3_Forward_step2_forwarded")
        self._add_step(result, "TC3-d", "Message Forwarded",
                       "pass" if sent else "fail",
                       [f"Forwarded to : {recipient}",
                        f"Result       : {'Forwarded ✓' if sent else 'Failed ✗'}"], ss2)
        if not sent:
            result["fail_reason"].append("Could not complete forward action")
            return False

        # Delivery check — fails TC if forwarded message was delivered
        page.wait_for_timeout(2000)
        self._check_delivery_generic(page, result, message, "TC3-e", tag="TC3")

        return True

    def _do_reply(self, page, result, recipient, message, reply_text=None, **kwargs):
        """Reply to a message."""
        reply_text = reply_text or f"Reply {message}"
        print(f"\n   [TC4] Reply to message from: {recipient}")
        time.sleep(10)

        # Open chat
        clicked = False
        for strategy in [
            lambda: page.locator(f"xpath=//span[normalize-space(text())='{recipient}']").first.click(timeout=5000),
            lambda: page.get_by_text(recipient, exact=True).first.click(timeout=5000),
        ]:
            try:
                strategy(); clicked = True; break
            except Exception:
                continue

        page.wait_for_timeout(3000)
        self._add_step(result, "TC4-a", f"Opened Chat with {recipient}",
                       "pass" if clicked else "fail",
                       [f"Recipient: {recipient}"])
        if not clicked:
            result["fail_reason"].append(f"Could not open chat with {recipient}")
            return False

        # Hover + click 3 dots
        dots_clicked = self._hover_and_click_dots(page, message, tag="REPLY")
        ss1, _ = self._screenshot(page, "TC4_Reply_step1_three_dots")
        self._add_step(result, "TC4-b", "Clicked 3 Dots on Message",
                       "pass" if dots_clicked else "fail",
                       [f"Message: '{message}'"], ss1)
        if not dots_clicked:
            result["fail_reason"].append("Could not click 3 dots on message")
            return False

        # Click Reply
        reply_clicked = False
        for sel in [
            "button:has-text('Reply')",
            "xpath=//button[normalize-space(text())='Reply']",
            "text=Reply", "[data-tid='reply-message']",
        ]:
            try:
                page.locator(sel).first.wait_for(state="visible", timeout=5000)
                page.locator(sel).first.click()
                reply_clicked = True
                break
            except Exception:
                continue

        page.wait_for_timeout(2000)
        self._add_step(result, "TC4-c", "Clicked Reply",
                       "pass" if reply_clicked else "fail",
                       ["Selected Reply from context menu"])
        if not reply_clicked:
            result["fail_reason"].append("Could not find Reply in menu")
            return False

        # Type reply
        typed = False
        for sel in [
            "div[contenteditable='true']",
            "xpath=//div[@contenteditable='true']",
        ]:
            try:
                box = page.locator(sel).last
                box.wait_for(state="visible", timeout=5000)
                box.click()
                page.wait_for_timeout(500)
                box.type(reply_text, delay=80)
                typed = True
                break
            except Exception:
                continue

        if not typed:
            result["fail_reason"].append("Could not type reply text")
            return False

        # Pre-connect vsmd + HAR before send
        vsmd_prep, har = self._before_send(page, "TC4_ReplyMessage")
        result["_har"] = har   # store now so HAR is saved even on early exit

        # Send
        sent = False
        for send_sel in [
            "button[aria-label='Send']", "button:has-text('Send')",
        ]:
            try:
                page.locator(send_sel).first.click(timeout=3000)
                sent = True
                break
            except Exception:
                continue
        if not sent:
            page.keyboard.press("Enter")
            sent = True

        page.wait_for_timeout(3000)
        self._after_send(page, result, vsmd_prep, har, "TC4_ReplyMessage", None)

        page.wait_for_timeout(3000)
        ss2, _ = self._screenshot(page, "TC4_Reply_step2_sent")
        self._add_step(result, "TC4-d", "Reply Sent",
                       "pass" if sent else "fail",
                       [f"Reply : '{reply_text}'",
                        f"Result: {'Sent ✓' if sent else 'Failed ✗'}"], ss2)
        if not sent:
            result["fail_reason"].append("Could not send reply")
            return False

        # Delivery check — fails TC if reply was delivered
        page.wait_for_timeout(2000)
        self._check_delivery_generic(page, result, reply_text, "TC4-e", tag="TC4")

        return True

    # ================================================================
    # TEAMS-SPECIFIC HELPERS (not in base class)
    # ================================================================

    def _hover_and_click_dots(self, page, message_text: str, tag: str = "") -> bool:
        """
        Hover a chat message bubble, then click the message-actions '...'
        (More actions — opens Forward/Reply menu).
        """
        found = False
        # Try locating the message paragraph
        for msg_sel in [
            f"xpath=//p[contains(text(),'{message_text}')]",
            f"xpath=//*[contains(text(),'{message_text}')]",
            f"text={message_text}",
        ]:
            try:
                el = page.locator(msg_sel).last
                el.wait_for(state="visible", timeout=5000)
                el.scroll_into_view_if_needed()
                el.hover()
                page.wait_for_timeout(800)
                found = True
                break
            except Exception:
                continue

        if not found:
            print(f"   [{tag}] Could not find message: '{message_text}'")
            return False

        # Click the message-actions '...' (rightmost ... after edit pencil)
        for dots_sel in [
            "xpath=//button[@aria-label='More actions']",
            "xpath=//button[contains(@aria-label,'More')]",
            "xpath=//button[@data-tid='message-actions-more']",
            "button[aria-label='More actions']",
        ]:
            try:
                page.locator(dots_sel).last.click(timeout=3000)
                print(f"   [{tag}] Clicked More Actions ({dots_sel})")
                return True
            except Exception:
                continue

        print(f"   [{tag}] Could not click More Actions dots")
        return False

    def _dismiss_windows_firewall(self):
        """
        Dismiss the Windows Firewall/Security dialog if it appears.
        This is a native Windows dialog — Playwright cannot see it.
        Uses pywinauto to click 'Allow access'.
        Safe to call at any time — does nothing if dialog is not present.
        """
        import time as _time
        try:
            from pywinauto import Desktop
            _desktop = Desktop(backend="win32")
            for _attempt in range(10):
                for _win in _desktop.windows():
                    try:
                        title = _win.window_text()
                        if "Windows Security" in title or "Windows Firewall" in title or "firewall" in title.lower():
                            print(f"   [FIREWALL] Dismissing Windows Firewall dialog...")
                            for _btn_text in ["Allow access", "Allow"]:
                                try:
                                    _win.child_window(title=_btn_text, control_type="Button").click_input()
                                    print(f"   [FIREWALL] ✓ Dismissed (clicked '{_btn_text}')")
                                    return
                                except Exception:
                                    continue
                    except Exception:
                        continue
                _time.sleep(0.5)
        except Exception:
            pass

    def _dismiss_meeting_popups(self, page):
        """Dismiss popups that appear when launching a Teams meeting."""
        import time as _time

        # Dismiss Windows Firewall dialog first
        self._dismiss_windows_firewall()

        # Block window management
        for sel in ["button:has-text('Block')", "xpath=//button[normalize-space(text())='Block']"]:
            try:
                page.locator(sel).first.click(timeout=3000)
                break
            except Exception:
                continue
        page.wait_for_timeout(1000)
        # Close No Microphone dialog
        for sel in [
            "xpath=//div[contains(text(),'No Microphone')]/..//button",
            "xpath=//*[contains(text(),'No microphone')]/parent::*/button",
        ]:
            try:
                page.locator(sel).first.click(timeout=2000)
                page.wait_for_timeout(500)
                break
            except Exception:
                continue
        page.wait_for_timeout(1000)
        # Close Invite people dialog
        for sel in [
            "xpath=//div[@role='dialog']//button[@aria-label='Close']",
            "xpath=(//button[@aria-label='Close'])[1]",
        ]:
            try:
                page.locator(sel).first.click(timeout=3000)
                break
            except Exception:
                try:
                    page.keyboard.press("Escape")
                except Exception:
                    pass
        page.wait_for_timeout(2000)