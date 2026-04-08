# Consilium Executive Coach

Voice-enabled executive coaching app with personal accounts, profile feedback loops, nudges, and an admin console.

## Run locally

1. Create `.env` in the app root:

```bash
cat > .env <<'EOF'
OPENAI_API_KEY=sk-REPLACE_WITH_REAL_KEY
OPENAI_MODEL=gpt-4o-mini
OPENAI_TTS_MODEL=gpt-4o-mini-tts
OPENAI_TTS_VOICE=shimmer
OPENAI_TTS_FORMAT=mp3
EOF
```

2. Start:

```bash
npm start
```

3. Open:

```text
http://127.0.0.1:8787
```

## First-time access

1. Click `Get Started`.
2. Sign in with default admin:
   - Username: `administrator`
   - Password: `password`
3. Change password and save profile settings.

## What’s included

- Natural conversational coaching with GROW + Kantor move guidance.
- High-fidelity TTS voice + browser fallback voice.
- Coachee profile updates after each interaction:
  - Big Five, MBTI, DISC
  - Strengths, development areas, score trends
  - Intervention tracking and outcomes
- Post-session feedback form (1-10 scores + comment) feeding back into profile/intervention updates.
- Nudge scheduling in settings with status flow:
  - `scheduled` -> `due` -> `completed` / `dismissed`
- Admin dashboard (shield icon):
  - All users, usage, use-cases, personality snapshots, feedback signals
  - Account controls per user: activate/deactivate, force reset, set temporary password

## Data and retention

- User/account data is stored in `data/users.json`.
- Session/profile/nudge/feedback retention is 365 days.
- Session cookies expire after 7 days.

## Publish publicly (Render)

1. Put this folder in a GitHub repo.
2. In Render, create a `Web Service` from that repo.
3. Set:
   - Runtime: `Node`
   - Build command: `npm install`
   - Start command: `npm start`
4. Add environment variables in Render:
   - `OPENAI_API_KEY`
   - `OPENAI_MODEL` (optional, default `gpt-4o-mini`)
   - `OPENAI_TTS_MODEL` (optional, default `gpt-4o-mini-tts`)
   - `OPENAI_TTS_VOICE` (optional, default `shimmer`)
   - `OPENAI_TTS_FORMAT` (optional, default `mp3`)
5. Add a persistent disk and mount it so `data/users.json` survives deploys. Without a persistent disk, accounts and history reset on restart.
6. Deploy, then share the Render URL.

## Notes

- The backend keeps API keys server-side.
- The server reads `.env` automatically when present.
- `HOST` defaults to `0.0.0.0`, so it works on hosted platforms.
