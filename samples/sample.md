---
aliases:
  - Test File
subtitle: MD to DOCX Test
---
# Heading Level 1

## Heading Level 2

### Heading Level 3

#### Heading Level 4

##### Heading Level 5

###### Heading Level 6

Setext Heading 1
================

Setext Heading 2
----------------

# Footnote Test

This text has a footnote[^1].

This text has another footnote[^note].

[^1]: This is footnote 1.
[^note]: This is a named footnote.

# Endnote Test

This text has an endnote[^endnote-1].

This text has another endnote[^endnote-ref].

[^endnote-1]: This is endnote 1, will appear at the end of document.
[^endnote-ref]: This is a named endnote.


# Heading with Text Test

## Title Ending with Punctuation:
When a heading ends with punctuation and is immediately followed by body text (no blank line), they form a combined heading-paragraph block.

## Heading without Symbol

This text is a new paragraph. There is a blank line above.

# Table Test

## Simple Table

| Name | Age |
| ---- | --- |
| Tom  | 5   |
| Bob  | 10  |
| Amy  | 15  |

## Empty Cell Table

| Food   | Price |
| ------ | ----- |
| Apple  | $1    |
|        | $2    |
| Banana | $3    |

# List Test

## Unordered List

- Cat
- Dog
- Bird

## Ordered List

1. Red
2. Blue
3. Green

## Nested List

- Animal
  - Cat
  - Dog
    - Big dog
    - Small dog
  - Bird
- Plant
  - Tree
  - Flower

## Mixed List

1. First item
   - Sub item A
   - Sub item B
2. Second item
   1. Sub item 1
   2. Sub item 2

# Quote Test

## Simple Quote

> This is a quote.
> It has two lines.

## Nested Quote (Compact Format)

> Level 1 quote.
>> Level 2 quote.
>>> Level 3 quote.

## Nested Quote (Spaced Format)

> Level 1 quote.
> > Level 2 quote.
> > > Level 3 quote.

# Code Test

## Code Block

```python
print("Hello")
x = 1 + 2
```

```javascript
console.log("Hi");
```

## Inline Code

This is `inline code` in text.

# Formula Test

## Inline Formula

The formula is $E = mc^2$ in this line.

## Block Formula

$$
a^2 + b^2 = c^2
$$

$$x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$$

# Text Style Test

This is **bold** text.

This is *italic* text.

This is ***bold and italic*** text.

This is ~~strikethrough~~ text.

# Link Test

This is a [link](https://example.com) in text.

# Horizontal Rule Test

Text before rule.

---

Text after dash rule.

***

Text after asterisk rule.

___

Text after underscore rule.
