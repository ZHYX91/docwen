# Multi-row three-line table borders

## Two header rows

| Region | Site | Scores | < |
| ^ | ^ | A | B |
| --- | --- || --- | --- |
| North | Alpha | 10 | < |
| ^ | Beta | ^ | ^ |

## Three header rows and nested groups

| Region | All | < | < | < |
| ^ | West | < | East | < |
| ^ | W1 | W2 | E1 | E2 |
| --- || --- | --- | --- | --- |
| North | 10 | < | 20 | 21 |
| ^ | ^ | ^ | 22 | 23 |
{repeat-header=true}

## Three header rows and a two-dimensional group

| Stub | Block | < | Tier |
| ^ | ^ | ^ | Sub |
| ^ | B1 | B2 | Leaf |
| --- || --- | --- | --- |
| North | 10 | < | 20 |
| ^ | ^ | ^ | 21 |
{repeat-header=false}
