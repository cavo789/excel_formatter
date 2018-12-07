<?php

// Only valid if PHP7 or greater
//declare(strict_types = 1);

/**
 * AUTHOR : AVONTURE Christophe.
 *
 * Written date : 29 october 2018
 *
 * Interface allowing to copy a long Excel formula and get a beautified version.
 * This version will make easier to read and understand the formula
 *
 * @see excelFormulaUtilitiesJS on https://github.com/joshbtn/excelFormulaUtilitiesJS
 */

define('REPO', 'https://github.com/cavo789/excel_formatter');

// Sample formula
$formula = '=IF(ISNA(VLOOKUP("Value";G1:K11;1;FALSE));"Not found";"Found")';

// Delimiter
$delim = ';';

// Get the GitHub corner
$github = '';
if (is_file($cat = __DIR__ . DIRECTORY_SEPARATOR . 'octocat.tmpl')) {
    $github = str_replace('%REPO%', REPO, file_get_contents($cat));
}
?>

<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8"/>
        <meta name="author" content="Christophe Avonture" />
        <meta name="robots" content="noindex, nofollow" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <meta http-equiv="content-type" content="text/html; charset=UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=9; IE=8;" />
        <title>Excel formula beautifier</title>
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/css/bootstrap.min.css">
        <link rel="stylesheet" href="assets/css/screen.css">
    </head>
    <body>
        <?php echo $github; ?>
        <div class="container">
            <div class="page-header"><h1>Excel formula beautifier</h1></div>
            <div class="container">
                <div class="form-group">
                    <label for="formula">Copy/Paste your Excel's formula in the 
                        textbox below then click on the Beautify button:</label>
                    <details>
                    <summary>How to use?</summary>
                    <div class="row">
                            <div class="col-sm">
                                <ul>
                                    <li>Type (or paste) the Excel formula to explain</li>
                                    <li>And click on the Beautify button.</li>
                                </ul>
                            </div>
                            <div class="col-sm">
                                <img src="https://raw.githubusercontent.com/cavo789/excel_formatter/master/images/demo.gif" alt="Demo">
                            </div>
                        </div>
                    </div>
                </details>
                    <textarea class="form-control" rows="3" id="formula" name="formula"><?php echo $formula; ?></textarea>
                </div>
                <div class="form-group row">
                    <label for="delim">Excel formula delimiter:</label>&nbsp;
                    <input type="text" style="width:30px;" size="3" value="<?php echo $delim; ?>" class="form-control" id="delim">
                </div>
                <button type="button" id="btnBeautify" class="btn btn-primary">Beautify</button>
                <hr/>
                <pre id="Result"></pre>
                <i style="display:block;font-size:0.6em;">
                    <a href="https://github.com/joshbtn/excelFormulaUtilitiesJS">
                        Excel beautifier written by Josh Bennett
                    </a>
                </i>
            </div>
        </div>
        <script src="https://code.jquery.com/jquery-3.3.1.min.js" integrity="sha256-FgpCb/KJQlLNfOu91ta32o/NMZxltwRo8QtmkMRdAu8=" crossorigin="anonymous"></script>
        <script type="text/javascript" src="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/js/bootstrap.min.js"></script>
        <script type="text/javascript" src="assets/js/excel-formula.min.js"></script>
        <script type="text/javascript">

            String.prototype.replaceAll = function(search, replacement) {
                var target = this;
                return target.replace(new RegExp(search, 'g'), replacement);
            };

            $('#btnBeautify').click(function(e)  {

                e.stopImmediatePropagation();

                var formula = $('#formula').val();
                if ($('#delim').val()==";") {
                    formula = formula.replaceAll(";", ",");
                }

                // Call the beautifier
                var formattedFormula = excelFormulaUtilities.formatFormulaHTML(formula);

                if ($('#delim').val()==";") {
                    formattedFormula = formattedFormula.replaceAll(",", ";");
                }
                
                // And output the result in the DOM element
                $('#Result').html(formattedFormula);
            });
        </script>
    </body>
</html>
