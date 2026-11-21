# TrialFence Pitch Deck — Speaker Notes

Keep the deck on screen; do not read it word-for-word. Aim for 4–5 minutes before the live demo.

## Slide 1 — One trial. One anonymous enrollment. (25 seconds)
TrialFence stops the same person from joining the same clinical trial protocol at multiple sites, without creating a shared identity database. The volunteer proves eligibility anonymously; the shared gate accepts the first enrollment and blocks the repeat.

## Slide 2 — The problem (35 seconds)
Clinical trials need clean enrollment data, but sites do not want to exchange raw participant IDs. Research shows that hidden participation and changed identity details are real sources of study contamination. Today, teams often choose between weak duplicate checks and risky data sharing.

## Slide 3 — The solution (35 seconds)
TrialFence changes what sites share. The coordinator still checks identity locally, but sites share only an anonymous cryptographic proof. The proof shows that the volunteer is eligible and has not already enrolled in this protocol. It does not reveal the volunteer’s identity.

## Slide 4 — How it works (45 seconds)
The browser creates a private credential. A coordinator verifies a fictional demo ID and adds only an anonymous commitment. The browser then makes a proof tied to the selected protocol and site. The smart contract accepts the first proof and blocks the same protocol-specific private tag at another site.

## Slide 5 — Working product (40 seconds)
This is a running full-stack prototype with volunteer, coordinator, and sponsor workspaces. In the tested flow, Site North was accepted, Site South was blocked as a duplicate with a signed denial record, and the same identity was accepted in a different protocol with an unrelated tag. The repository includes 22 contract tests and a full integration suite.

## Slide 6 — Impact and pilot (30 seconds)
Sponsors get cleaner trial data, sites get an instant accept-or-block answer, and IRBs get an audit trail without raw cross-site identity. The next practical step is a small pilot: one protocol, two or three sites, privacy review, onboarding, and testnet validation.

## Slide 7 — Live demo and close (20 seconds, then demo)
We will now register a protocol and two sites, issue anonymous membership, accept the first enrollment, block the duplicate, and show that a new protocol remains unlinkable. Close with: “TrialFence protects trial integrity without building a participant identity database.”
