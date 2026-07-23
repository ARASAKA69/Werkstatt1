# Queue Page Overrides

> **PROJECT:** Lager Kisten Klaerung
> **Generated:** 2026-07-23 16:57:40
> **Page Type:** Product Detail

> ⚠️ **IMPORTANT:** Rules in this file **override** the Master file (`design-system/MASTER.md`).
> Only deviations from the Master are documented here. For all other rules, refer to the Master.

---

## Page-Specific Rules

### Layout Overrides

- **Max Width:** 1200px (standard)
- **Layout:** Full-width sections, centered content

### Spacing Overrides

- No overrides — use Master spacing

### Typography Overrides

- No overrides — use Master typography

### Color Overrides

Style is **Modern Dark (Cinema)** — remap Master teal/orange onto dark surfaces:

| Role | Hex |
|------|-----|
| Background deep/base | `#0a0a0f` → `#050506` |
| Elevated / glass | `#0e1014` + `rgba(255,255,255,0.05)` |
| Foreground | `#EDEDEF` / muted `#8A8F98` |
| Primary / ring | `#0D9488` / secondary `#14B8A6` |
| Accent CTA | `#EA580C` |
| Border | `rgba(255,255,255,0.08)` |

### Component Overrides

- Glass sidebar + header (`backdrop-filter: blur(16px)`)
- Dense queue list, 40px min touch targets
- Skeleton loaders while queue/detail fetch
- Inline SVG icons (Heroicons-style), no emoji
- Press scale `0.97`, easing `cubic-bezier(0.16, 1, 0.3, 1)`
- Ambient teal/orange blobs (respect `prefers-reduced-motion`)

---

## Page-Specific Components

- No unique components for this page

---

## Recommendations

- Effects: Product animation playback, step progression animations, hover reveal effects, smooth zoom on interaction
