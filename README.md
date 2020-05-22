# DataTables Buttons Html5 Excel Styling

**Add beautifully styled Excel output to your DataTables.**

DataTables is an amazing tool to display your tables in a user friendly way, and the buttons extension makes downloading those tables a breeze. Now you can simply style your downloaded tables without having to learn the intricacies of Open Office XML Styling.

## Demo

[View the Excel styling demo](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/simple_table_style.html)

## Usage

1. Include jQuery

2. Include Datatables css and js (https://www.datatables.net/download/)

3. Include the plugins' style and optional template javascript files

```html
<script src="js/buttons.html5.styles.js"></script>
<script src="js/buttons.html5.styles.templates.js"></script>
```

4. Add an excelStyles config option to apply a style or template to the Excel file. [View demo](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/single_template_style.html)

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
