# Datatables Buttons Html5 Excel Styling

## Demo

[View the Excel style demo](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/simple_table_style.html)

## Usage

Include jQuery:

Include Datatables css and js (https://www.datatables.net/download/)

Include the plugins' style and optional template javascript files from cdn or download this repo and include the files in the js folder.

``` html
<script src="https://cdn.jsdelivr.net/npm/datatables-buttons-excel-styles@0.7.1/js/buttons.html5.styles.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/datatables-buttons-excel-styles@0.7.1/js/buttons.html5.styles.templates.min.js"></script>
```

Add an excelStyles config option to apply a style or template when you download your DataTable with the Excel button.

``` html

$('#myTable').DataTable( {
    dom: 'Bfrtip',
    buttons: [
      {
        extend: 'excel',
        excelStyles: {
          template: 'blue_medium'
        }
      }
    ]
} );
```

** More docs underway - due end of May 2020  **
