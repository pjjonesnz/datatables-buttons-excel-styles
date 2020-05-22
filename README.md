# Datatables Buttons Html5 Excel Styling

## Demo

[View the Excel style demo](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/simple_table_style.html)

## Usage

Include jQuery

Include Datatables css and js (https://www.datatables.net/download/)

Include the plugins' style and optional template javascript files

```html
<script src="js/buttons.html5.styles.js"></script>
<script src="js/buttons.html5.styles.templates.js"></script>
```

Add an excelStyles config option to apply a style or template when you download your DataTable with the Excel button. [View this example](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/single_template_style.html)

```js
$('#myTable').DataTable({
    dom: 'Bfrtip',
    buttons: [
        {
            extend: 'excel',              // Extend the excel button
            excelStyles: {                // Add an excelStyles definition
                template: 'blue_medium',  // Apply the 'blue_medium' template
            },
        },
    ],
});
```

** More docs underway - due end of May 2020 **
