---
name: PosterAgain Design System
colors:
  surface: '#F8FAFB'
  surface-dim: '#dbdad9'
  surface-bright: '#fbf9f8'
  surface-container-lowest: '#ffffff'
  surface-container-low: '#F0F4F6'
  surface-container: '#efeded'
  surface-container-high: '#e9e8e7'
  surface-container-highest: '#e4e2e2'
  on-surface: '#1b1c1c'
  on-surface-variant: '#3f484d'
  inverse-surface: '#303030'
  inverse-on-surface: '#f2f0f0'
  outline: '#70787e'
  outline-variant: '#bfc8cd'
  surface-tint: '#016684'
  primary: '#004d64'
  on-primary: '#FFFFFF'
  primary-container: '#BEE9FF'
  on-primary-container: '#a2e1ff'
  inverse-primary: '#87d0f2'
  secondary: '#4d616c'
  on-secondary: '#ffffff'
  secondary-container: '#d0e6f3'
  on-secondary-container: '#536772'
  tertiary: '#434365'
  on-tertiary: '#ffffff'
  tertiary-container: '#5b5b7e'
  on-tertiary-container: '#d6d4fe'
  error: '#BA1A1A'
  on-error: '#ffffff'
  error-container: '#ffdad6'
  on-error-container: '#93000a'
  primary-fixed: '#bee9ff'
  primary-fixed-dim: '#87d0f2'
  on-primary-fixed: '#001f2a'
  on-primary-fixed-variant: '#004d64'
  secondary-fixed: '#d0e6f3'
  secondary-fixed-dim: '#b4cad6'
  on-secondary-fixed: '#081e27'
  on-secondary-fixed-variant: '#364954'
  tertiary-fixed: '#e1dfff'
  tertiary-fixed-dim: '#c4c3eb'
  on-tertiary-fixed: '#181837'
  on-tertiary-fixed-variant: '#434465'
  background: '#fbf9f8'
  on-background: '#1b1c1c'
  surface-variant: '#e4e2e2'
  success: '#006D39'
typography:
  display-lg:
    fontFamily: Noto Sans
    fontSize: 57px
    fontWeight: '400'
    lineHeight: 64px
    letterSpacing: -0.25px
  headline-lg:
    fontFamily: Noto Sans
    fontSize: 32px
    fontWeight: '400'
    lineHeight: 40px
  headline-lg-mobile:
    fontFamily: Noto Sans
    fontSize: 28px
    fontWeight: '400'
    lineHeight: 36px
  title-lg:
    fontFamily: Noto Sans
    fontSize: 22px
    fontWeight: '500'
    lineHeight: 28px
  body-lg:
    fontFamily: Noto Sans
    fontSize: 16px
    fontWeight: '400'
    lineHeight: 24px
    letterSpacing: 0.5px
  body-md:
    fontFamily: Noto Sans
    fontSize: 14px
    fontWeight: '400'
    lineHeight: 20px
    letterSpacing: 0.25px
  label-lg:
    fontFamily: Noto Sans
    fontSize: 14px
    fontWeight: '500'
    lineHeight: 20px
    letterSpacing: 0.1px
  label-sm:
    fontFamily: Noto Sans
    fontSize: 11px
    fontWeight: '500'
    lineHeight: 16px
    letterSpacing: 0.5px
rounded:
  sm: 0.25rem
  DEFAULT: 0.5rem
  md: 0.75rem
  lg: 1rem
  xl: 1.5rem
  full: 9999px
spacing:
  margin-mobile: 16px
  margin-tablet: 24px
  margin-desktop: 32px
  gutter: 16px
  baseline: 4px
---

## Brand & Style
This design system is engineered for educators and administrators, prioritizing clarity, efficiency, and a sense of calm. Drawing heavily from **Material 3 (M3)** and **Material You** principles, the system leverages tonal palettes and dynamic color logic to reduce cognitive load during high-stakes administrative tasks.

The visual style is **Modern Corporate / Professional**, utilizing expansive whitespace and a structured "Surface-Container" architecture. By following Material You’s tonal relationship rules, the interface feels integrated and personalized, ensuring that the "efficiency tool" aspect remains gentle rather than aggressive.

## Colors
The palette is rooted in a calming office aesthetic, moving away from harsh pure blacks and high-vibrancy primaries toward **soft teals and muted blues**. 

- **Primary:** A professional teal used for the main action buttons and active states.
- **Surface/Container:** Used for the card-based layout to differentiate content areas without heavy borders.
- **Semantic Logic:** 
    - **Success (Green):** Specifically reserved for matched data fields or successful data imports.
    - **Error (Red):** Used for unmatched alerts or critical missing information.
    - **Tonal Palettes:** All secondary and tertiary UI elements should be derived from the primary hue's tonal palette (e.g., a Primary 90 for containers) to maintain a harmonious, low-contrast environment.

## Typography
The system utilizes **Noto Sans** (including SC for localized contexts) to ensure maximum legibility and a neutral, professional tone. 

The type hierarchy follows M3's scale. Use `Title Large` for app bars and `Body Large` for primary content to ensure the interface is accessible for users who may be glancing at their devices in a busy classroom or office. `Label` styles should be used for button text and helper text within input fields to maintain a clear distinction from content.

## Layout & Spacing
This system employs a **Fluid Grid** model based on a 4dp baseline grid. 

- **Mobile:** 4-column layout with 16dp side margins.
- **Tablet/Foldable:** 12-column layout with 24dp margins, often utilizing a navigation rail instead of a bottom bar.
- **Desktop:** 12-column layout with a maximum content width of 1200px, centered on the screen.

Spacing between related elements (like a label and its input) should use 8dp (2x baseline), while spacing between sections should use 24dp or 32dp to create a clear "gentle" breathing room.

## Elevation & Depth
Elevation is communicated through **Tonal Layers** rather than heavy shadows, following Material 3's modern approach. 

- **Level 0 (Flat):** The main background surface.
- **Level 1 (Elevated):** Standard cards and list items. Use a subtle color shift (Surface Container Low) and a very soft, diffused shadow (0dp to 1dp equivalent).
- **Level 2+:** Reserved for interactive elements like Floating Action Buttons (FABs) or Dialogs, which use a slightly more pronounced shadow and a lighter tonal tint to appear closer to the user.

## Shapes
In line with the "gentle" brand persona, the system uses **Rounded** geometry. 

- **Small Components (Chips, Tooltips):** 8px corner radius.
- **Medium Components (Cards, Dialogs):** 12px to 16px corner radius.
- **Large Components (Buttons, FABs):** Fully rounded (pill-shaped) or 28px radius to signify high interactivity.

## Components
- **Buttons:** Use M3 "Filled" buttons for primary actions (e.g., "Post Grade") and "Tonal" buttons for secondary actions. All buttons feature large, accessible touch targets.
- **Cards:** Elevated cards serve as the primary container for student records or administrative tasks. Use a 1px subtle stroke in a neutral tone if additional definition is needed on Level 0 surfaces.
- **Bottom Navigation:** Standard M3 bar with 3-5 destinations, using active indicator "pills" to highlight the current section.
- **List Tiles:** Feature a leading icon (rounded) and trailing metadata or chevron. Use `Body Medium` for supporting text.
- **Input Fields:** Filled text fields with a bottom indicator line; the fill color should be a slightly darker shade of the surface to guide the eye.
- **Matched Field Indicator:** Use a Success-toned chip or highlight within list tiles to provide immediate visual confirmation of matched data.