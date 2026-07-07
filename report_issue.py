"""Report-an-issue section: lets a user email a problem PDF for investigation.

A mailto: link can pre-fill the recipient, subject, and body, but browsers
and email clients deliberately block JS/HTML from attaching files to it (that
would be a trivial way to smuggle attachments without consent) — so the PDF
itself can never be attached automatically. The UI is built around that
constraint: it opens the user's email client fully pre-filled, and explicitly
tells them to drag the PDF in themselves before sending.
"""

import urllib.parse

import streamlit as st

# The real address lives in .streamlit/secrets.toml (gitignored) so it never
# lands in git history — see .streamlit/secrets.toml.example for the format.
SUPPORT_EMAIL = st.secrets.get("support_email", "")
SUBJECT = "PDF to EXCEL issue"
BODY = (
    "Hi,\n\n"
    "The attached PDF produced a wrong/unexpected result in the LCC PDF to "
    "Excel tool.\n\n"
    "What went wrong: <describe here — e.g. missing column, wrong date, "
    "wrong row count>\n\n"
    "(Please attach the problem PDF before sending.)"
)


def render_report_issue() -> None:
    """Render a single-line 'report a bad PDF' footer with a mailto link.

    The link label deliberately doesn't print the raw address as visible
    text — it only appears inside the mailto: href, which keeps it off the
    rendered page for casual viewing (though, like any client-side link,
    it's still visible to anyone who inspects the page's HTML source).
    """
    if not SUPPORT_EMAIL:
        return  # secrets.toml missing/misconfigured — fail silently, no crash

    mailto = (
        f"mailto:{SUPPORT_EMAIL}"
        f"?subject={urllib.parse.quote(SUBJECT)}"
        f"&body={urllib.parse.quote(BODY)}"
    )
    st.caption(f"Found a PDF that gives a wrong result? [Report it]({mailto}), attach the PDF before sending.")
