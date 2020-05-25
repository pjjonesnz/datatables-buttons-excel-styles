## Patterns

(https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_patternFill_topic_ID0E6KM6.html)

## Gradients

(https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_gradientFill_topic_ID0ENWD6.html)

### Linear Gradient

```js
"excelStyles":
  {
    "cells": "3:n,2",
    "style": {
      "fill": {
        "gradient": {
          "degree": "90",
          "stop": [
            {
              "position": "0",
              "color": "FF92D050"
            },
            {
              "position": "1",
              "color": "FF0000"
            }
           ]
        }
      }
    }
  }
```

### Path Gradient

```js
"excelStyles":
  {
    "cells": "3:n,2",
    "style": {
      "fill": {
        "gradient": {
          "type": "path",
          "left": "0.2",
          "right": "0.8",
          "top": "0.2",
          "bottom": "0.8",
          "stop": [
            {
              "position": "0",
              "color": "FF92D050"
            },
            {
              "position": "1",
              "color": "FF0000"
            }
          ]
        }
      }
    }
  }
```
