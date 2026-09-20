# 3 propositions — bouton « Remplacer »

Contexte repris du prompt : bouton actuel = rectangle bordé teal, fond blanc, 13px, à droite de « Encadrant : Prénom Nom ». Palette : teal `#0f6e6a` / `#0f5754`, ambre focus `#e3a331`, glassmorphism, coins arrondis 14-18px, Space Grotesk + Manrope.

---

## Proposition A — Chip pilule avec pastille icône

**Idée** : un bouton en forme de pilule (arrondi complet), avec l'icône ⇄ isolée dans une petite pastille pleine teal, et le libellé à côté. À l'état actif (remplaçant choisi), la pilule s'inverse en fond teal plein.

- **Simplicité** : reste un bouton texte + icône, donc l'action est aussi explicite qu'aujourd'hui — aucune perte de clarté, juste une forme plus soignée.
- **Harmonie visuelle** : la forme pilule et la pastille pleine font écho aux badges/accordéons déjà présents (`accordion-toggle`) et à la rondeur générale de l'UI (cartes 14-18px), au lieu du rectangle actuel qui tranche.
- **Responsive** : identique mobile/desktop (la pilule est compacte), donc aucun risque de rupture de mise en page en ligne à côté du texte encadrant.

```html
<button type="button" id="remplacantOuvrir" class="replacement-chip" aria-haspopup="dialog">
  <span class="replacement-chip-icon" aria-hidden="true">⇄</span>
  <span class="replacement-chip-label">Remplacer</span>
</button>
```

```css
.replacement-chip {
  display: inline-flex; align-items: center; gap: 6px;
  min-height: 44px; padding: 6px 14px 6px 6px; flex-shrink: 0;
  background: rgba(15,110,106,0.10);
  border: 1px solid rgba(15,110,106,0.25);
  border-radius: 999px;
  color: #0f5754; font: inherit; font-size: 13px; font-weight: 600;
  cursor: pointer; transition: background .15s ease, border-color .15s ease;
}
.replacement-chip-icon {
  display: inline-flex; align-items: center; justify-content: center;
  width: 24px; height: 24px; border-radius: 50%;
  background: #0f6e6a; color: #fff; font-size: 13px;
}
.replacement-chip:hover { background: rgba(15,110,106,0.18); }
.replacement-chip:focus-visible { outline: 3px solid #e3a331; outline-offset: 2px; }
.replacement-chip.is-active { background: #0f6e6a; border-color: #0f6e6a; color: #fff; }
.replacement-chip.is-active .replacement-chip-icon { background: rgba(255,255,255,0.25); }
```

- **Avantage** : le plus sûr des trois — aussi lisible qu'aujourd'hui, juste mieux intégré.
- **Limite** : gain visuel réel mais assez sage, ne change pas fondamentalement l'expérience.

---

## Proposition B — Icône seule sur mobile, icône + libellé sur desktop

**Idée** : un bouton carré compact (44×44px) avec juste l'icône ⇄ sur mobile (là où l'espace horizontal est précieux), qui s'élargit pour afficher le libellé à côté de l'icône dès que l'écran est plus large (desktop). Une pastille ambre signale un remplacement actif.

- **Simplicité** : l'icône seule peut être ambiguë pour un utilisateur peu à l'aise → compensée par `aria-label`/`title` et par la pastille de statut, mais reste le point faible de cette piste.
- **Harmonie visuelle** : très discret, laisse le texte « Encadrant : Prénom Nom » respirer, cohérent avec le style épuré des cartes.
- **Responsive** : c'est la piste qui exploite le plus le responsive — comportement réellement différent mobile/desktop, pensé pour économiser l'espace là où c'est utile.

```html
<button type="button" id="remplacantOuvrir" class="replacement-icon-btn" aria-haspopup="dialog" aria-label="Remplacer l'encadrant" title="Remplacer l'encadrant">
  <span aria-hidden="true">⇄</span>
  <span class="replacement-icon-btn-label">Remplacer</span>
  <span class="replacement-dot" hidden aria-hidden="true"></span>
</button>
```

```css
.replacement-icon-btn {
  position: relative; display: inline-flex; align-items: center; gap: 6px;
  min-width: 44px; height: 44px; padding: 0 10px; flex-shrink: 0;
  background: rgba(255,255,255,0.6);
  border: 1px solid rgba(15,110,106,0.35);
  border-radius: 12px; color: #0f6e6a; font-size: 18px; cursor: pointer;
}
.replacement-icon-btn-label { display: none; font-size: 13px; font-weight: 600; }
.replacement-icon-btn:hover { background: rgba(15,110,106,0.12); }
.replacement-icon-btn:focus-visible { outline: 3px solid #e3a331; outline-offset: 2px; }
.replacement-dot {
  position: absolute; top: 4px; right: 4px; width: 8px; height: 8px;
  border-radius: 50%; background: #e3a331;
}
@media (min-width: 601px) {
  .replacement-icon-btn-label { display: inline; }
}
```

- **Avantage** : le plus « propre » sur mobile, où chaque pixel compte.
- **Limite** : demande de bien mettre à jour `aria-label`/`title` en JS selon l'état (« Remplacer » ↔ « Modifier »), sinon perte de clarté.

---

## Proposition C — Lien discret intégré au texte

**Idée** : abandonner l'aspect « bouton » et traiter « Remplacer » comme un lien souligné, à la suite du texte « Encadrant : Prénom Nom », dans la même police et taille — comme une action secondaire inline plutôt qu'un contrôle séparé.

- **Simplicité** : plus ambigu qu'un bouton — un lien souligné est une convention connue, mais moins « tapable » visuellement qu'un bouton avec fond, ce qui peut réduire la découvrabilité sur mobile.
- **Harmonie visuelle** : la plus intégrée des trois — aucun encart, se fond dans la ligne de texte, esthétique très épurée.
- **Responsive** : simple à faire tenir sur une ligne qui wrap naturellement en cas de nom long, mais la zone de tap doit être élargie en padding invisible pour respecter les 44px sans casser l'alignement visuel.

```html
<span class="encadrant-info">Encadrant : <strong>Edward LICHTNER</strong></span>
<button type="button" id="remplacantOuvrir" class="replacement-link" aria-haspopup="dialog">
  <span aria-hidden="true">⇄</span> Remplacer
</button>
```

```css
.replacement-link {
  display: inline-flex; align-items: center; gap: 4px;
  min-height: 44px; padding: 4px 6px; flex-shrink: 0;
  background: none; border: none;
  color: #0f6e6a; font: inherit; font-size: 13px; font-weight: 600;
  text-decoration: underline; text-decoration-color: rgba(15,110,106,0.4);
  text-underline-offset: 3px; cursor: pointer;
}
.replacement-link:hover { text-decoration-color: #0f6e6a; }
.replacement-link:focus-visible { outline: 3px solid #e3a331; outline-offset: 2px; border-radius: 4px; }
.replacement-link.is-active { color: #a15b00; text-decoration-color: rgba(161,91,0,0.4); }
```

- **Avantage** : le plus léger visuellement, cohérent avec une UI qui se veut discrète.
- **Limite** : le moins évident à repérer au premier coup d'œil — risque pour un encadrant pressé qui ne « voit » pas l'action.

---

## Recommandation

**Proposition A (chip pilule)** en priorité : c'est celle qui améliore le plus l'harmonie visuelle sans rien sacrifier à la simplicité d'usage — le bouton reste aussi explicite qu'aujourd'hui, juste visuellement intégré. La proposition B est intéressante si l'économie d'espace mobile devient un vrai enjeu (ligne encadrant très chargée), mais elle demande plus de rigueur (aria-label dynamique). La proposition C est la plus élégante visuellement mais la plus risquée en terrain (utilisateurs pressés, parfois mains humides) : à réserver si l'app vise un public déjà très habitué.
