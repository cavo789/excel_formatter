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
 *
 * Last mod:
 * 2018-12-31 - Abandonment of jQuery and migration to vue.js
 */

define('REPO', 'https://github.com/cavo789/excel_formatter');

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
            <div class="container" id="app">
                <div class="form-group">
                    <how-to-use demo="https://raw.githubusercontent.com/cavo789/excel_formatter/master/images/demo.gif">
                        <ul>
                            <li>Type (or paste) the Excel formula to explain</li>
                            <li>And click on the Beautify button.</li>
                        </ul>
                    </how-to-use>
                    <label for="formula">Copy/Paste your Excel's formula in the
                        textbox below then click on the Beautify button:</label>
                    <textarea class="form-control" rows="3" name="formula" v-model="formula"></textarea>
                </div>
                <div class="form-group row">
                    <label for="delim">Excel formula delimiter:</label>&nbsp;
                    <input type="text" style="width:30px;" size="3" v-model="delim" class="form-control">
                </div>
                <button type="button" class="btn btn-primary" @click="processBeautify">Beautify</button>
                <hr/>
                <pre v-html="HTML"></span>
                <i style="display:block;font-size:0.6em;">
                    <a href="https://github.com/joshbtn/excelFormulaUtilitiesJS">
                        Excel beautifier written by Josh Bennett
                    </a>
                </i>
            </div>
        </div>
        <script src="https://unpkg.com/vue"></script>
        <script type="text/javascript" src="assets/js/excel-formula.min.js"></script>
        <script type="text/javascript">

            String.prototype.replaceAll = function(search, replacement) {
                var target = this;
                return target.replace(new RegExp(search, 'g'), replacement);
            };

            Vue.component('how-to-use', {
                props: {
                    demo: {
                        type: String,
                        required: true
                    }
                },
                template:
                    `<details>
                        <summary>How to use?</summary>
                        <div class="row">
                                <div class="col-sm">
                                    <slot></slot>
                                </div>
                                <div class="col-sm"><img v-bind:src="demo" alt="Demo"></div>
                            </div>
                        </div>
                    </details>`
            });

            var app = new Vue({
                el: '#app',
                data: {
                    formula: '=IF(ISNA(VLOOKUP("Value";G1:K11;1;FALSE));"Not found";"Found")',
                    delim: ';',
                    HTML: ''
                },
                methods: {
                    processBeautify() {

                        this.formula = this.formula.trim();

                        if (this.delim == ";") {
                            this.formula = this.formula.replaceAll(";", ",");
                        }

                        // Call the beautifier
                        var formattedFormula = excelFormulaUtilities.formatFormulaHTML(this.formula);

                        if (this.delim == ";") {
                            formattedFormula = formattedFormula.replaceAll(",", ";");
                        }

                        // And output the result in the DOM element
                        this.HTML = formattedFormula;
                    }
                }
            });
        </script>
    </body>
</html>
