# Flat Allocation System

## What is this?

When an old building is torn down and a new one is built in its place, every flat owner from the old building needs to be given a flat in the new building. This app does that allocation fairly and randomly, without human bias, or picking favourites.

You upload a file describing the old building, the new building, and the list of flat owners. The app assigns each owner a new flat and gives you the results as a downloadable spreadsheet.

## Why does this exist?

In most housing societies today, flat allocation is done by drawing chits from a bowl. While that feels random, it's hard to verify and easy to question.

This app replaces chits with a transparent, repeatable algorithm. If given the same random seed number, the allocation will always produce the exact same result, so that anyone can independently verify that the process was fair. The seed acts like a public receipt of the randomisation.

## What goes in the input file?

The input is an Excel file (`.xlsx`) with these sheets:

| Sheet | What it contains |
|---|---|
| **Old Building** | Each row is a floor in the old building - wing name, floor number, and how many flats are on that floor. |
| **New Building** | Same format but for the new building. The new building must have at least as many flats as the old one. |
| **Flat Owners** | A numbered list of every flat owner - their flat number, name, contact info, and (optionally) a constraint ID if they have a special requirement. |
| **Constraints** *(optional)* | Special requirements - see "Preferences & Groups" below. |
| **Randomisation Seed** *(optional)* | A fixed number in cell B5. If provided, the app uses this number to drive the randomisation so the result is reproducible. If left empty, the app picks a random seed each time. |

A sample input file is included in the app for reference.

## What comes out?

The output is an Excel file with five sheets:

| Sheet | What it contains |
|---|---|
| **Allocation** | The main result - each old flat number mapped to its new flat (wing, floor, unit), along with how it was allocated. |
| **Validation** | A checklist confirming every constraint was satisfied and no flat was assigned twice. |
| **Building Layout** | A visual grid of the new building showing which old flat owner ended up where. |
| **Audit Trail** | A step-by-step log of every allocation decision the algorithm made, in order. |
| **Metadata** | Summary info - timestamp, seed used, building sizes, and overall status. |

## The Custom Script Feature

Click **"Edit Script"** in the app to open a panel that shows the full source code of the allocation logic. You can copy it, paste it into an AI tool (like Claude or similar), describe what you want to change, and paste the modified script back. The app will use your custom version instead of the default.

This means **anyone can modify how the allocation works without being a programmer**. This feature is purely to speed up the developemnt process and try things out fast, without needing a deployment.

Hit "Reset Script" at any time to go back to the default.

## How the Algorithm Works

The algorithm runs in three phases, from most constrained to least:

1. **Groups first.** If flat owners need to be neighbours (e.g. a family across multiple flats), the algorithm finds spots where they can be placed in consecutive adjacent units within the same wing, overflowing to consecutive floors above when needed. Larger groups are handled before smaller groups because they're harder to fit.

2. **Preferences next.** If an owner has a preference for a specific wing, floor, or unit, the algorithm picks a random available flat that matches. More specific preferences (e.g. wing + floor + unit) are handled before vague ones (e.g. just a wing).

3. **Everyone else randomly.** All remaining owners are shuffled and each one gets a randomly chosen flat from whatever's still available.

At every step, the algorithm uses a seeded random number generator - meaning the "randomness" is locked to the seed number and fully reproducible.

## Preferences & Groups

There are two types of special requirements you can set in the Constraints sheet:

### Group
Any number of flat owners (2 or more) who must be placed near each other in the same wing.
- Members are placed in **consecutive adjacent units**, filling one floor first.
- If a floor can't hold all members, the remainder overflows to the **next consecutive floor above**, again in consecutive units.
- This continues across as many consecutive floors as needed.
- You can optionally restrict a group to a specific wing or starting floor.

### Preference
One or more flat owners who want a specific wing, floor, or unit number - or any combination of the three. The more specific the request, the earlier it gets handled (so it's more likely to be satisfied).