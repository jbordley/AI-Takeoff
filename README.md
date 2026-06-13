# AI Takeoff Canvas

[![Live demo](https://img.shields.io/badge/live%20demo-jbordley.github.io-2ea44f?logo=github)](https://jbordley.github.io/AI-Takeoff/)

An AI-powered AV/low-voltage system design tool built for integrators.
Sits between lightweight tools like System Surveyor and enterprise platforms
like D-Tools — faster than one, smarter than the other.

> **▶ Try it live:** **https://jbordley.github.io/AI-Takeoff/**
> Runs entirely in your browser — no install, no backend. Ships with a
> synthetic illustrative product library; bring your own pricing data to use
> it for real.

## What It Does

- Canvas-based drawing with 30 zone types (Conference Room, Bedroom, Theater, etc.)
- Auto-populate AV devices based on zone type and room rules
- Color-coded zones with matching device symbols
- BOM generation with split client/dealer export
- PDF and DOCX proposal output
- PDF floorplan import with calibration and scale tools
- AI agent layer (in development) — vision-based device placement recommendations

## Status

MVP complete. AI agent integration in active development.

## Tech Stack

- HTML5 / JavaScript / Fabric.js canvas
- OpenAI GPT-4 Vision (agent layer)
- Pluggable product/pricing library (bring your own data)
- Distributor pricing API integration (planned)

## Product Library

This repo ships a small **synthetic, illustrative** sample library
(`data/snapone-library.json`) so the app is demoable out of the box. The product
names, SKUs, and prices in it are placeholders. Swap in your own catalog and
pricing — the app reads the same schema (`products[]` with `dealerPrice`, `msrp`,
`margin`, `laborHours`, `laborRate`, and per-tier `tierPricing`).

> Note: distributor/dealer pricing is typically confidential under your dealer
> agreement. Do not commit real dealer pricing to a public repository.

## Roadmap

- [ ] AI zone analysis — GPT-4 Vision reads floorplans, places devices
- [ ] NLP chat interface for conversational design refinement
- [ ] Distributor pricing API integration
- [ ] Rules engine with AVIXA standards-based placement logic
- [ ] ID plan symbol recognition — read existing TV locations from drawings
