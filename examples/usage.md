# Usage examples

Tool-call argument shapes for the two creation modes. (Base64 values are abbreviated.)

## From scratch (neutral theme)

`pptx_create_deck`:

```json
{
  "title": "Q3 Review",
  "slides": [
    { "archetype": "cover", "title": "Q3 Review", "eyebrow": "FY26", "subtitle": "Results and outlook" },
    { "archetype": "agenda", "title": "Agenda", "items": [
      { "label": "Results", "duration": "10m" },
      { "label": "Pipeline", "duration": "15m" }
    ]},
    { "archetype": "statsBanner", "title": "Highlights", "stats": [
      { "value": "+18%", "label": "revenue", "trend": "YoY" },
      { "value": "0", "label": "churn" }
    ]},
    { "archetype": "closingNextSteps", "title": "Next steps", "steps": [
      { "title": "Approve budget", "owner": "Finance" }
    ]}
  ]
}
```

## Bring your own template

Step 1 - discover the template's stencil slides:

`pptx_list_layouts`:

```json
{ "template": { "base64": "UEsDBBQ...=" } }
```

returns, among the archetype catalog, something like:

```json
{
  "templateSlides": [
    { "index": 0, "layoutName": "Title Slide", "placeholderTypes": ["ctrTitle", "subTitle"] },
    { "index": 1, "layoutName": "Title and Content", "placeholderTypes": ["title", "body"] }
  ]
}
```

Step 2 - build new slides cloned from those stencils:

`pptx_create_deck`:

```json
{
  "template": { "base64": "UEsDBBQ...=" },
  "slides": [
    { "archetype": "cover", "stencilSlideIndex": 0, "title": "Project Atlas", "subtitle": "Kickoff" },
    { "archetype": "contentBullets", "stencilSlideIndex": 1, "title": "Goals", "bullets": ["Ship v1", "Onboard 3 teams"] },
    { "archetype": "comparisonTable", "stencilSlideIndex": 1, "title": "Options",
      "options": ["Build", "Buy"], "features": [{ "name": "Cost", "values": ["$$", "$"] }] }
  ]
}
```

## Validate / audit / repair any deck

```json
{ "source": { "base64": "UEsDBBQ...=" } }
```

`pptx_validate` -> pass/fail + violations. `pptx_audit` -> that plus metrics and portability risk. `pptx_repair` -> a fixed file + a change log.
