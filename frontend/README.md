## frontend (next.js ui)

next.js app for timelinecraft. provides landing, operations, and settings pages with a shadcn‑style sidebar for app pages.

### features
- minimal landing at `/landing` with demo preview and cta
- operations at `/operate`
  - step tabs: 1) set tags, 2) choose params, 3) preview
  - draggable tags list with stable ids
  - sliders for sidebar width and item height; morph toggle; file input
  - preview section (36 dummy slides as placeholders)
- settings at `/settings` with theme toggle (defaults to dark, persisted)
- sidebar appears on app pages via `app/(app)/layout.tsx`

### tech
- next.js 16 (app router), react 19
- tailwind css v4
- shadcn‑style sidebar (`components/ui/sidebar.tsx`, `components/app-sidebar.tsx`) – see docs: https://ui.shadcn.com/docs/components/sidebar

### scripts
```bash
npm run dev     # start local dev server
npm run build   # production build
npm start       # start production build
npm run lint    # run linter
```

### development
- entry pages:
  - `/landing`  – marketing/demo (no sidebar)
  - `/operate`  – main workflow (with sidebar)
  - `/settings` – theme preferences (with sidebar)
- utilities:
  - `lib/utils.ts` – `cn()` class merge helper
  - `lib/indexes.ts` – `IndexItem` type, ID creation, reorder helper

### docker (optional)
from repo root you can use compose targets. example (if configured):
```bash
docker-compose up -d
docker-compose logs -f frontend
```


