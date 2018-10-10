declare module GC{
    module Spread{
        module CalcEngine{
            /**
             * Represents the expression type
             * @enum {number}
             */
            export enum ExpressionType{
                /**
                 * Specifies the unknow type
                 */
                unknow= 0,
                /**
                 * Specifies the reference type
                 */
                reference= 1,
                /**
                 * Specifies the number type
                 */
                number= 2,
                /**
                 * Specifies the string type
                 */
                string= 3,
                /**
                 * Specifies the boolean type
                 */
                boolean= 4,
                /**
                 * Specifies the error type
                 */
                error= 5,
                /**
                 * Specifies the array type
                 */
                array= 6,
                /**
                 * Specifies the function type
                 */
                function= 7,
                /**
                 * Specifies the name type
                 */
                name= 8,
                /**
                 * Specifies the operator type
                 */
                operator= 9,
                /**
                 * Specifies the parenthesses type
                 */
                parentheses= 10,
                /**
                 * Specifies the missing argument type
                 */
                missingArgument= 11,
                /**
                 * Specifies the expand type
                 */
                expand= 12,
                /**
                 * Specifies the struct reference type
                 */
                structReference= 13
            }


            export class AsyncEvaluateContext{
                /**
                 * Represents an evaluate context for async functions.
                 * @class
                 * @param {GC.Spread.CalcEngine.EvaluateContext} context The common evaluate context.
                 */
                constructor();
                /**
                 * Set the async function evaluate result to CalcEngine, CalcEngine uses this value to recalculate the cell that contains this async function and all dependents cells.
                 * @param {object} value The async function evaluate result.
                 */
                setAsyncResult(value:  any): void;
            }

            export class CalcError{
                /**
                 * Represents an error in calculation.
                 * @class
                 * @param {string} error The description of the error.
                 */
                constructor(error:  string);
                /**
                 * Parses the specified error from the string.
                 * @param {string} value The error string.
                 * @returns {GC.Spread.CalcEngine.CalcError} The calculation error.
                 */
                static parse(value: string): GC.Spread.CalcEngine.CalcError;
                /**
                 * Returns a string that represents this instance.
                 * @returns {string} The error string.
                 */
                toString(): string;
            }

            export class Expression{
                /**
                 * Provides the base class from which the classes that represent expression tree nodes are derived. This is an abstract class.
                 * @class
                 */
                constructor(type:  ExpressionType);
                /**
                 * Indicates the expression type
                 * @type {GC.Spread.CalcEngine.ExpressionType}
                 */
                type: ExpressionType;
            }
            module Functions{
                /**
                 * Defines a global custom function.
                 * @param {string} name The name of the function.
                 * @param {GC.Spread.CalcEngine.Functions.Function} fn The function to add.
                 */
                function defineGlobalCustomFunction(name:  string,  fn:  Functions.Function): void;
                /**
                 * Finds a function by name.
                 * @param {string} name The name of the function.
                 * @returns {GC.Spread.CalcEngine.Functions.Function} The function with the specified name.
                 */
                function findGlobalFunction(name:  string): Functions.Function;

                export interface IFunctionDescription{
                    description: string;
                    parameters: GC.Spread.CalcEngine.Functions.IParameterDescription[];
                }


                export interface IParameterDescription{
                    name: string;
                    repeatable?: boolean;
                    optional?: boolean;
                }

                /**
                 * Represents the asynchronous function evaluate mode.
                 */
                export enum AsyncFunctionEvaluateMode{
                    /**
                     * Specifies that the async function evaluates the changed, referenced cells.
                     */
                    onRecalculation= 0,
                    /**
                     * Specifies that the async function only evaluates once.
                     */
                    calculateOnce= 1,
                    /**
                     * Specifies that the async function evaluates based on the interval.
                     */
                    onInterval= 2
                }


                export class AsyncFunction{
                    /**
                     * Represents an abstract base class for defining asynchronization functions.
                     * @class
                     * @param {string} name The name of the function.
                     * @param {number} minArgs The minimum number of arguments for the function.
                     * @param {number} maxArgs The maximum number of arguments for the function.
                     * @param {object} description The description of the function.
                     */
                    constructor(name:  string,  minArgs:  number,  maxArgs:  number,  description:  Object);
                    /**
                     * Returns the default value of the evaluated function result before getting the async result.
                     * @returns {object} The default value of the evaluated function result before getting the async result.
                     */
                    defaultValue(): any;
                    /**
                     * Returns the result of the function applied to the arguments.
                     * @param {GC.Spread.CalcEngine.AsyncEvaluateContext} context The evaluate context
                     * @param {object} args Arguments for the function evaluation
                     * @returns {object} The result of the function applied to the arguments.
                     */
                    evaluateAsync(context:  AsyncEvaluateContext,  args:  any[]): any;
                    /**
                     * Decides how to re-calculate the formula.
                     * @returns {GC.Spread.CalcEngine.Functions.AsyncFunctionEvaluateMode} The evaluate mode.
                     */
                    evaluateMode(): any;
                    /**
                     * Returns the interval.
                     * @returns {number} The interval in milliseconds.
                     */
                    interval(): any;
                }

                export class Function{
                    /**
                     * Represents an abstract base class for defining functions.
                     * @class
                     * @param {string} name The name of the function.
                     * @param {number} minArgs The minimum number of arguments for the function.
                     * @param {number} maxArgs The maximum number of arguments for the function.
                     * @param {object} functionDescription The description object of the function.<br />
                     ** functionDescription.description {string} The description of the function.<br />
                     ** functionDescription.parameters {ParameterDescription[]} The parameters' description array of the function.<br />
                     *** ParameterDescription {object} The parameter's description object.<br />
                     *** ParameterDescription.name {string} The parameter name.<br />
                     *** ParameterDescription.repeatable {boolean} Whether the parameter is repeatable.<br />
                     *** ParameterDescription.optional {boolean} Whether the parameter is optional.
                     */
                    constructor(name?:  string,  minArgs?:  number,  maxArgs?:  number,  description?:  GC.Spread.CalcEngine.Functions.IFunctionDescription);
                    /**
                     * Represents the maximum number of arguments for the function.
                     * @type {number}
                     */
                    maxArgs: number;
                    /**
                     * Represents the minimum number of arguments for the function.
                     * @type {number}
                     */
                    minArgs: number;
                    /**
                     * Represents the name of the function.
                     * @type {string}
                     */
                    name: string;
                    /**
                     * Represents the type name string used for supporting serialization.
                     * @type {string}
                     */
                    typeName: string;
                    /**
                     * Determines whether the function accepts array values for the specified argument.
                     * @param {number} argIndex Index of the argument.
                     * @function
                     * @returns {boolean} <c>true</c> if the function accepts array values for the specified argument; otherwise, <c>false</c>.
                     */
                    acceptsArray(argIndex:  number): boolean;
                    /**
                     * Indicates whether the function can process Error values.
                     * @param {number} argIndex Index of the argument.
                     * @returns {boolean} <c>true</c> if the function can process Error values for the specified argument; otherwise, <c>false</c>.
                     * @function
                     */
                    acceptsError(argIndex:  number): boolean;
                    /**
                     * Indicates whether the Evaluate method can process missing arguments.
                     * @param {number} argIndex Index of the argument
                     * @returns {boolean} <c>true</c> if the Evaluate method can process missing arguments; otherwise, <c>false</c>.
                     */
                    acceptsMissingArgument(argIndex:  number): boolean;
                    /**
                     * Determines whether the function accepts Reference values for the specified argument.
                     * @param {number} argIndex Index of the argument.
                     * @returns {boolean} <c>true</c> if the function accepts Reference values for the specified argument; otherwise, <c>false</c>.
                     * @function
                     */
                    acceptsReference(argIndex:  number): boolean;
                    /**
                     * Returns the description of the function.
                     * @function
                     * @returns {object} The description of the function.
                     */
                    description(): Object;
                    /**
                     * Returns the result of the function applied to the arguments.
                     * @param {object} args Arguments for the function evaluation
                     * @returns {object} The result of the function applied to the arguments.
                     */
                    evaluate(args:  any): any;
                    /**
                     * Finds the branch argument.
                     * @param {object} test The test.
                     * @returns {number} Indicates the index of the argument that would be treated as the branch condition.
                     */
                    findBranchArgument(test:  any): number;
                    /**
                     * Finds the test argument when this function is branched.
                     * @returns {number} Indicates the index of the argument that would be treated as the test condition.
                     */
                    findTestArgument(): boolean;
                    /**
                     * Gets a value that indicates whether this function is branched by arguments as conditional.
                     * @returns {boolean} <c>true</c> if this instance is branched; otherwise, <c>false</c>.
                     */
                    isBranch(): boolean;
                    /**
                     * Determines whether the evaluation of the function is dependent on the context in which the evaluation occurs.
                     * @returns {boolean} <c>true</c> if the evaluation of the function is dependent on the context; otherwise, <c>false</c>.
                     */
                    isContextSensitive(): boolean;
                    /**
                     * Determines whether the function is volatile while it is being evaluated.
                     * @returns {boolean} <c>true</c> if the function is volatile; otherwise, <c>false</c>.
                     */
                    isVolatile(): boolean;
                }
            }

        }

        module Commands{
            /**
             * Represents the key code.
             * @enum {number}
             */
            export enum Key{
                /**
                 * Indicates the left arrow key.
                 */
                left= 37,
                /**
                 * Indicates the right arrow key.
                 */
                right= 39,
                /**
                 * Indicates the up arrow key.
                 */
                up= 38,
                /**
                 * Indicates the down arrow key.
                 */
                down= 40,
                /**
                 * Indicates the Tab key.
                 */
                tab= 9,
                /**
                 * Indicates the Enter key.
                 */
                enter= 13,
                /**
                 * Indicates the Shift key.
                 */
                shift= 16,
                /**
                 * Indicates the Ctrl key.
                 */
                ctrl= 17,
                /**
                 * Indicates the space key.
                 */
                space= 32,
                /**
                 * Indicates the Alt key.
                 */
                altkey= 18,
                /**
                 * Indicates the Home key.
                 */
                home= 36,
                /**
                 * Indicates the End key.
                 */
                end= 35,
                /**
                 * Indicates the Page Up key.
                 */
                pup= 33,
                /**
                 * Indicates the Page Down key.
                 */
                pdn= 34,
                /**
                 * Indicates the Backspace key.
                 */
                backspace= 8,
                /**
                 * Indicates the Delete key.
                 */
                del= 46,
                /**
                 * Indicates the Esc key.
                 */
                esc= 27,
                /**
                 * Indicates the A key
                 */
                a= 65,
                /**
                 * Indicates the C key.
                 */
                c= 67,
                /**
                 * Indicates the V key.
                 */
                v= 86,
                /**
                 * Indicates the X key.
                 */
                x= 88,
                /**
                 * Indicates the Z key.
                 */
                z= 90,
                /**
                 * Indicates the Y key.
                 */
                y= 89
            }


            export class CommandManager{
                /**
                 * Represents a command manager.
                 * @class
                 * @param {Object} context The execution context for all commands in the command manager.
                 * @constructor
                 */
                constructor(context:  Object);
                /**
                 * Executes a command and adds the command to UndoManager.
                 * @param {Object} commandOptions The options for the command.<br />
                 ** commandOptions.cmd {string} The command name, the field is required.<br />
                 ** commandOptions.arg1 {Object} The command argument, the name and type are decided by the command definition.<br />
                 ** commandOptions.arg2 {Object} The command argument, the name and type are decided by the command definition.<br />
                 ** commandOptions.argN {Object} The command argument, the name and type are decided by the command definition.<br />
                 **
                 ** For example, the following code executes the autoFitColumn command.<br />
                 ** var spread = GC.Spread.Sheets.findControl(document.getElementById("ss"));<br />
                 ** spread.commandManager().execute({cmd: "autoFitColumn", sheetName: "Sheet1", columns: [{col: 1}], rowHeader: false, autoFitType: GC.Spread.Sheets.AutoFitType.cell});
                 * @returns {string} The execute command result.
                 */
                execute(commandOptions:  Object): any;
                /**
                 * Registers a command with the command manager.
                 * @param {string} name The name of the command.
                 * @param {Object} command The object that defines the command.
                 ** For example, the following code registers the changeBackColor command and then executes the command.<br />
                 ** var command = {<br />
                 **     canUndo: true,<br />
                 ** 	execute: function (context, options, isUndo) {<br />
                 ** 	    var sheet = context.getSheetFromName(options.sheetName);<br />
                 ** 		var cell = sheet.getCell(options.row, options.col);<br />
                 ** 		if (isUndo) {<br />
                 ** 		    cell.backColor(options._oldBackColor);<br />
                 ** 		} else {<br />
                 ** 			options._oldBackColor = cell.backColor();<br />
                 ** 		    cell.backColor(options.backColor);<br />
                 ** 		}<br />
                 ** 	}<br />
                 ** };<br />
                 ** var spread = GC.Spread.Sheets.findControl(document.getElementById("ss"));<br />
                 ** var commandManager = spread.commandManager();<br />
                 ** commandManager.register("changeBackColor", command);<br />
                 ** commandManager.execute({cmd: "changeBackColor", sheetName: spread.getSheet(0).name(), row: 1, col: 2, backColor: "red"});
                 * @param {number|GC.Spread.Commands.Key} key The key code.
                 * @param {boolean} ctrl <c>true</c> if the command uses the Ctrl key; otherwise, <c>false</c>.
                 * @param {boolean} shift <c>true</c> if the command uses the Shift key; otherwise, <c>false</c>.
                 * @param {boolean} alt <c>true</c> if the command uses the Alt key; otherwise, <c>false</c>.
                 * @param {boolean} meta <c>true</c> if the command uses the Command key on the Macintosh or the Windows key on Microsoft Windows; otherwise, <c>false</c>.
                 */
                register(name:  string,  command:  Object,  key?:  number|GC.Spread.Commands.Key,  ctrl?:  boolean,  shift?:  boolean,  alt?:  boolean,  meta?:  boolean): void;
                /**
                 * Binds a shortcut key to a command.
                 * @param {string} commandName The command name, setting commandName to undefined removes the bound command of the shortcut key.
                 * @param {number|GC.Spread.Commands.Key} key The key code, setting the key code to undefined removes the shortcut key of the command.
                 * @param {boolean} ctrl <c>true</c> if the command uses the Ctrl key; otherwise, <c>false</c>.
                 * @param {boolean} shift <c>true</c> if the command uses the Shift key; otherwise, <c>false</c>.
                 * @param {boolean} alt <c>true</c> if the command uses the Alt key; otherwise, <c>false</c>.
                 * @param {boolean} meta <c>true</c> if the command uses the Command key on the Macintosh or the Windows key on Microsoft Windows; otherwise, <c>false</c>.
                 */
                setShortcutKey(commandName:  string,  key?:  number|GC.Spread.Commands.Key,  ctrl?:  boolean,  shift?:  boolean,  alt?:  boolean,  meta?:  boolean): void;
            }

            export class UndoManager{
                /**
                 * Represents the undo manager.
                 * @constructor
                 */
                constructor();
                /**
                 * Gets whether the redo operation is allowed.
                 * @returns {boolean} <c>true</c> if the redo operation is allowed; otherwise, <c>false</c>.
                 */
                canRedo(): boolean;
                /**
                 * Gets whether the undo operation is allowed.
                 * @returns {boolean} <c>true</c> if the undo operation is allowed; otherwise, <c>false</c>.
                 */
                canUndo(): boolean;
                /**
                 * Clears all of the undo stack and the redo stack.
                 */
                clear(): void;
                /**
                 * Redoes the last command.
                 * @returns {boolean} <c>true</c> if the redo operation is successful; otherwise, <c>false</c>.
                 */
                redo(): boolean;
                /**
                 * Undoes the last command.
                 * @returns {boolean} <c>true</c> if the undo operation is successful; otherwise, <c>false</c>.
                 */
                undo(): boolean;
            }
        }

        module Common{

            export interface IDateTimeFormat{
                abbreviatedDayNames?: string[];
                abbreviatedMonthGenitiveNames?: string[];
                abbreviatedMonthNames?: string[];
                amDesignator?: string;
                dayNames?: string[];
                fullDateTimePattern?: string;
                longDatePattern?: string;
                longTimePattern?: string;
                monthDayPattern?: string;
                monthGenitiveNames?: string[];
                monthNames?: string[];
                pmDesignator?: string;
                shortDatePattern?: string;
                shortTimePattern?: string;
                yearMonthPattern?: string;
            }


            export interface INumberFormat{
                currencyDecimalSeparator?: string;
                currencyGroupSeparator?: string;
                currencySymbol?: string;
                numberDecimalSeparator?: string;
                numberGroupSeparator?: string;
                listSeparator?: string;
                arrayListSeparator?: string;
                arrayGroupSeparator?: string;
                dbNumber?: Object
            }


            export class CultureInfo{
                /**
                 * Represents the custom culture class. The member variable can be overwritten.
                 * @class
                 */
                constructor();
                /**
                 * Indicates the date time format fields.
                 * @type {Object}
                 * @property {Array.<string>} abbreviatedDayNames - Specifies the day formatter for "ddd".
                 * @property {Array.<string>} abbreviatedMonthGenitiveNames - Specifies the month formatter for "MMM".
                 * @property {Array.<string>} abbreviatedMonthNames - Specifies the month formatter for "MMM".
                 * @property {string} amDesignator - Indicates the AM designator.
                 * @property {Array.<string>} dayNames - Specifies the day formatter for "dddd".
                 * @property {string} fullDateTimePattern - Specifies the standard date formatter for "F".
                 * @property {string} longDatePattern - Specifies the standard date formatter for "D".
                 * @property {string} longTimePattern - Specifies the standard date formatter for "T" and "U".
                 * @property {string} monthDayPattern - Specifies the standard date formatter for "M" and "m".
                 * @property {Array.<string>} monthGenitiveNames - Specifies the formatter for "MMMM".
                 * @property {Array.<string>} monthNames - Specifies the formatter for "M" or "MM".
                 * @property {string} pmDesignator - Indicates the PM designator.
                 * @property {string} shortDatePattern - Specifies the standard date formatter for "d".
                 * @property {string} shortTimePattern - Specifies the standard date formatter for "t".
                 * @property {string} yearMonthPattern - Specifies the standard date formatter for "y" and "Y".
                 */
                DateTimeFormat: GC.Spread.Common.IDateTimeFormat;
                /**
                 * Indicates all the number format fields.
                 * @type {Object}
                 * @property {string} currencyDecimalSeparator - Indicates the currency decimal point.
                 * @property {string} currencyGroupSeparator - Indicates the currency thousand separator.
                 * @property {string} currencySymbol - Indicates the currency symbol.
                 * @property {string} numberDecimalSeparator - Indicates the decimal point.
                 * @property {string} numberGroupSeparator - Indicates the thousand separator.
                 * @property {string} listSeparator - Indicates the separator for function arguments in a formula.
                 * @property {string} arrayListSeparator - Indicates the separator for the constants in one row of an array constant in a formula.
                 * @property {string} arrayGroupSeparator - Indicates the separator for the array rows of an array constant in a formula.
                 * @property {object} dbNumber - Specifies the DBNumber characters.
                 * The dbNumber object structure:<br />
                 *  {
                 *     1: {letters: ['\u5146', '\u5343', '\u767e', '\u5341', '\u4ebf', '\u5343', '\u767e', '\u5341', '\u4e07', '\u5343', '\u767e', '\u5341', ''], // \u5146\u5343\u767e\u5341\u4ebf\u5343\u767e\u5341\u4e07\u5343\u767e\u5341
                 *         numbers: ['\u25cb', '\u4e00', '\u4e8c', '\u4e09', '\u56db', '\u4e94', '\u516d', '\u4e03', '\u516b', '\u4e5d'] }, // \u25cb\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d
                 *     2: {letters: ['\u5146', '\u4edf', '\u4f70', '\u62fe', '\u4ebf', '\u4edf', '\u4f70', '\u62fe', '\u4e07', '\u4edf', '\u4f70', '\u62fe', ''], // \u5146\u4edf\u4f70\u62fe\u4ebf\u4edf\u4f70\u62fe\u4e07\u4edf\u4f70\u62fe
                 *         numbers: ['\u96f6', '\u58f9', '\u8d30', '\u53c1', '\u8086', '\u4f0d', '\u9646', '\u67d2', '\u634c', '\u7396']}, // \u96f6\u58f9\u8d30\u53c1\u8086\u4f0d\u9646\u67d2\u634c\u7396
                 *     3: {letters: null,
                 *         numbers: ['\uff10', '\uff11', '\uff12', '\uff13', '\uff14', '\uff15', '\uff16', '\uff17', '\uff18', '\uff19']} // \uff10\uff11\uff12\uff13\uff14\uff15\uff16\uff17\uff18\uff19
                 * };
                 */
                NumberFormat: GC.Spread.Common.INumberFormat;
            }

            export class CultureManager{
                /**
                 * Represents a culture manager.
                 * @constructor
                 */
                constructor();
                /**
                 * Adds the cultureInfo into the culture manager.
                 * @static
                 * @param {string} cultureName
                 * @param {GC.Spread.Common.CultureInfo} culture object
                 */
                static addCultureInfo(cultureName: any,  culture: any): GC.Spread.Common.CultureInfo;
                /**
                 * Get or set the Sheets culture.
                 * @static
                 * @param {string} cultureName The culture name to get.
                 * @returns {string} The current culture name of Sheets.
                 */
                static culture(cultureName?:  string): string;
                /**
                 * Gets the specified cultureInfo. If no culture name, get current cultureInfo.
                 * @static
                 * @param {Object} cultureName Culture name or culture ID
                 * @returns {GC.Spread.Common.CultureInfo} cultureInfo object
                 */
                static getCultureInfo(cultureName:  Object): GC.Spread.Common.CultureInfo;
            }
        }

        module Excel{

            export class IO{
                /**
                 * Represents an excel import and export class.
                 * @class
                 */
                constructor();
                /**
                 * Imports an excel file.
                 * @param {Blob} file The excel file.
                 * @param {function} successCallBack Call this function after successfully loading the file. function (json) { }.
                 * @param {function} errorCallBack Call this function if an error occurs. The exception parameter object structure { errorCode: GC.Spread.Excel.IO.ErrorCode, errorMessage: string}.
                 * @param {Object} options The options for import excel.<br />
                 ** options.password {string} the excel file's password.
                 */
                open(file:  Blob,  successCallBack:  Function,  errorCallBack:  Function,  options?:  any): void;
                /**
                 * Creates and saves an excel file with the SpreadJS json.
                 * @param {object} json The spread sheets json object, or string.
                 * @param {function} successCallBack Call this function after successfully exporting the file. function (blob) { }.
                 * @param {function} errorCallBack Call this function if an error occurs. The exception parameter object structure { errorCode: GC.Spread.Excel.IO.ErrorCode, errorMessage: string}.
                 * @param {Object} options The options for export excel.<br />
                 ** options.password {string} the excel file's password.
                 */
                save(json:  String,  successCallBack:  Function,  errorCallBack:  Function,  options?:  any): void;
            }
            module IO{
                /**
                 * Specifies the excel io error code.
                 * @enum {number}
                 */
                export enum ErrorCode{
                    /**
                     *  File read and write exception.
                     */
                    fileIOError= 0,
                    /**
                     *  Incorrect file format.
                     */
                    fileFormatError= 1,
                    /**
                     *  The Excel file cannot be opened because the workbook/worksheet is password protected.
                     */
                    noPassword= 2,
                    /**
                     *  The specified password is incorrect.
                     */
                    invalidPassword= 3
                }

            }

        }

        module Formatter{

            export class FormatterBase{
                /**
                 * Represents a custom formatter with the specified format string.
                 * @class
                 * @param {string} format The format.
                 * @param {string} cultureName The culture name.
                 */
                constructor(format:  string,  cultureName:  string);
                /** Represents the type name string used for supporting serialization.
                 * @type {string}
                 */
                typeName: string;
                /**
                 * Formats the specified object as a string with a conditional color. This function should be overwritten.
                 * @param {Object} obj The object with cell data to format.
                 * @returns {string} The formatted string.
                 */
                format(obj:  Object): string;
                /**
                 * Loads the object state from the specified JSON string.
                 * @param {Object} settings The custom formatter data from deserialization.
                 */
                fromJSON(settings:  Object): void;
                /**
                 * Parses the specified text. This function should be overwritten.
                 * @param {string} text The text.
                 * @returns {Object} The parsed object.
                 */
                parse(str:  string): Object;
                /**
                 * Saves the object state to a JSON string.
                 * @returns {Object} The custom formatter data.
                 */
                toJSON(): Object;
            }

            export class GeneralFormatter{
                /**
                 * Represents a formatter with the specified format mode and format string.
                 * @class
                 * @param {string} format The format.
                 * @param {string} cultureName The culture name.
                 */
                constructor(format:  string,  cultureName:  string);
                /**
                 * Formats the specified object as a string with a conditional color.
                 * @param {Object} obj The object with cell data to format.
                 * @returns {string} The formatted string.
                 */
                format(obj:  Object): string;
                /**
                 * Gets or sets the format string for this formatter.
                 * @param {string} value The format string for this formatter.
                 * @returns {string|GC.Spread.Formatter.GeneralFormatter} If no value is set, returns the formatter string for this formatter; otherwise, returns the formatter.
                 */
                formatString(value?:  string): any;
                /**
                 * Parses the specified text.
                 * @param {string} text The text.
                 * @returns {Object} The parsed object.
                 */
                parse(str:  string): Object;
            }
        }

        module Sheets{
            /**
             * Represents the license key for evaluation version and production version.
             */
            var LicenseKey: string;
            /**
             * Gets the Workbook instance by the host element.
             * @param {HTMLElement|string} host The host element or the host element id.
             * @returns {GC.Spread.Sheets.Workbook} The Workbook instance.
             */
            function findControl(host:  HTMLElement): GC.Spread.Sheets.Workbook;
            /**
             * Gets the type from the type string. This method is used for supporting the serialization of the custom object.
             * @param {string} typeString The type string.
             * @returns {Object} The type.
             */
            function getTypeFromString(typeString:  string): any;

            export interface FloatingObjectLoadedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                floatingObject: FloatingObjects.FloatingObject;
                element: HTMLElement;
            }


            export interface IActiveSheetChangedEventArgs{
                oldSheet: GC.Spread.Sheets.Worksheet;
                newSheet: GC.Spread.Sheets.Worksheet;
            }


            export interface IActiveSheetChangingEventArgs{
                oldSheet: GC.Spread.Sheets.Worksheet;
                newSheet: GC.Spread.Sheets.Worksheet;
                cancel: boolean;
            }


            export interface IButtonClickedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                row: number;
                col: number;
                sheetArea: SheetArea;
            }


            export interface ICellChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                row: number;
                col: number;
                sheetArea: SheetArea;
                propertyName: string;
                oldValue: any;
                newValue: any;
            }


            export interface ICellClickEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                sheetArea: SheetArea;
                row: number;
                col: number;
            }


            export interface ICellDoubleClickEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                sheetArea: SheetArea;
                row: number;
                col: number;
            }


            export interface ICellPosition{
                row: number;
                col: number;
            }


            export interface IClipboardChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                copyData: string;
            }


            export interface IClipboardChangingEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                copyData: string;
                cancel: boolean;
            }


            export interface IClipboardPastedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                cellRange: Range;
                pasteOption: ClipboardPasteOptions;
            }


            export interface IClipboardPastingEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                cellRange: Range;
                pasteOption: ClipboardPasteOptions;
                cancel: boolean;
            }


            export interface IColumnChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                col: number;
                sheetArea: SheetArea;
                propertyName: string;
                oldValue: any;
                newValue: any;
                count?: number;
            }


            export interface IColumnWidthChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                colList: any[];
                header: boolean;
            }


            export interface IColumnWidthChangingEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                colList: any[];
                header: boolean;
                cancel: boolean;
            }


            export interface ICommentChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                comment: Comments.Comment;
                propertyName: string;
            }


            export interface ICommentRemovedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                comment: Comments.Comment;
            }


            export interface ICommentRemovingEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                comment: Comments.Comment;
                cancel: boolean;
            }


            export interface IDirtyCellInfo{
                row: number;
                col: number;
                newValue: any;
                oldValue: any;
            }


            export interface IDragDropBlockCompletedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                fromRow: number;
                fromCol: number;
                toRow: number;
                toCol: number;
                rowCount: number;
                colCount: number;
                copy: boolean;
                insert: boolean;
                copyOption: CopyToOptions;
            }


            export interface IDragDropBlockEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                fromRow: number;
                fromCol: number;
                toRow: number;
                toCol: number;
                rowCount: number;
                colCount: number;
                copy: boolean;
                insert: boolean;
                copyOption: CopyToOptions;
                cancel: boolean;
            }


            export interface IDragFillBlockCompletedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                fillRange: GC.Spread.Sheets.Range;
                autoFillType: GC.Spread.Sheets.Fill.AutoFillType;
                fillDirection: GC.Spread.Sheets.Fill.FillDirection;
            }


            export interface IDragFillBlockEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                fillRange: GC.Spread.Sheets.Range;
                autoFillType: GC.Spread.Sheets.Fill.AutoFillType;
                fillDirection: GC.Spread.Sheets.Fill.FillDirection;
                cancel: boolean;
            }


            export interface IEditChangeEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                row: number;
                col: number;
                editingText: any;
            }


            export interface IEditEndedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                row: number;
                col: number;
                editingText: Object;
            }


            export interface IEditEndingEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                row: number;
                col: number;
                editor: Object;
                editingText: Object;
                cancel: boolean;
            }


            export interface IEditorStatusChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                oldStatus: EditorStatus;
                newStatus: EditorStatus;
            }


            export interface IEditStartingEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                row: number;
                col: number;
                cancel: boolean;
            }


            export interface IEnterCellEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                row: number;
                col: number;
            }


            export interface IFilterButtonHitInfo{
                rowFilter: any;
                row?: number;
                col?: number;
                x?: number;
                y?: number;
                width?: number;
                height?: number;
                sheetArea?: GC.Spread.Sheets.SheetArea;
            }


            export interface IFloatingObjectChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                floatingObject: FloatingObjects.FloatingObject;
                propertyName: string;
            }


            export interface IFloatingObjectRemovedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                floatingObject: FloatingObjects.FloatingObject;
            }


            export interface IFloatingObjectRemovingEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                floatingObject: FloatingObjects.FloatingObject;
                cancel: boolean;
            }


            export interface IFloatingObjectSelectionChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                floatingObject: FloatingObjects.FloatingObject;
            }


            export interface IFormulaInformation{
                hasFormula: boolean;
                formula: string;
                isArrayFormula: boolean;
                baseRange: Range;
            }


            export interface IFormulaRangeHitInfo{
                paramRange: GC.Spread.Sheets.IParamRange; //param range info
                inTopLeft?: boolean;
                inTopRight?: boolean;
                inBottomLeft?: boolean;
                inBottomRight?: boolean;
                inBorder?: boolean;
            }


            export interface IHitTestCellTypeHitInfo{
                x?: number;
                y?: number;
                row?: number;
                col?: number;
                cellRect?: GC.Spread.Sheets.Rect;
                sheetArea?: GC.Spread.Sheets.SheetArea;
                isReservedLocation?: boolean;
            }


            export interface IHitTestCommentHitInfo{
                x?: number;
                y?: number;
                comment?: GC.Spread.Sheets.Comments.Comment;
                area?: string;
            }


            export interface IHitTestDragInfo{
                action?: string;
                side?: string;
                outside?: boolean;
            }


            export interface IHitTestFloatingObjectHitInfo{
                x?: number;
                y?: number;
                floatingObject?: GC.Spread.Sheets.FloatingObjects.FloatingObject;
            }


            export interface IHitTestInformation{
                x?: number;
                y?: number;
                rowViewportIndex?: number;
                colViewportIndex?: number;
                row?: number;
                col?: number;
                hitTestType?: GC.Spread.Sheets.SheetArea;
                resizeInfo?: GC.Spread.Sheets.IHitTestResizeInfo;
                outlineHitInfo?: GC.Spread.Sheets.IHitTestOutlineHitInfo;
                filterButtonHitInfo?: GC.Spread.Sheets.IFilterButtonHitInfo;
                dragInfo?: GC.Spread.Sheets.IHitTestDragInfo;
                cellTypeHitInfo?: GC.Spread.Sheets.IHitTestCellTypeHitInfo;
                floatingObjectHitInfo?: GC.Spread.Sheets.IHitTestFloatingObjectHitInfo;
                formulaRangeHitInfo?: GC.Spread.Sheets.IFormulaRangeHitInfo;
                commentHitInfo?: GC.Spread.Sheets.IHitTestCommentHitInfo;
            }


            export interface IHitTestOutlineHitInfo{
                what?: string;
                info?: GC.Spread.Sheets.IOutlineHitInfo;
            }


            export interface IHitTestResizeInfo{
                action?: string;
                index?: number;
                sheetArea?: GC.Spread.Sheets.SheetArea;
                startY?: number;
                movingY?: number;
                startX?: number;
                movingX?: number;
            }


            export interface IInvalidOperationEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                invalidType: InvalidOperationType;
                message: string;
            }


            export interface ILeaveCellEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                row: number;
                col: number;
                cancel: boolean;
            }


            export interface ILeftColumnChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                oldLeftCol: number;
                newLeftCol: number;
            }


            export interface IOutlineHitInfo{
                index?: number;
                isExpanded?: boolean;
                level?: number;
                lineDirection?: GC.Spread.Sheets.Outlines.OutlineDirection;
                paintLine?: boolean;
            }


            export interface IParamRange{
                textOffset: number; // range text offset in formulatextbox's value
                text: string; // range text
                index: number; // index in all ranges
            }


            export interface IPictureChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                picture: FloatingObjects.Picture;
                propertyName: string;
            }


            export interface IPictureSelectionChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                picture: FloatingObjects.Picture;
            }


            export interface IProtectionOptions{
                allowSelectLockedCells: boolean; //True or undefined if the user can select locked cells.
                allowSelectUnlockedCells: boolean; //True or undefined if the user can select unlocked cells.
                allowSort: boolean; //True if the user can sort ranges.
                allowFilter: boolean; //True if the user can filter ranges.
                allowEditObjects: boolean; //True if the user can edit floating objects.
                allowResizeRows: boolean; //True if the user can resize rows.
                allowResizeColumns: boolean; //True if the user can resize columns.
            }


            export interface IRangeChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                row: number;
                col: number;
                rowCount: number;
                colCount: number;
                changedCells: ICellPosition[];
                action: RangeChangedAction;
            }


            export interface IRangeFilteredEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                col: number;
                filterValues: any[];
            }


            export interface IRangeFilteringEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                col: number;
                filterValues: any[];
            }


            export interface IRangeGroupStateChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                isRowGroup: boolean;
                index: number;
                level: number;
            }


            export interface IRangeGroupStateChangingEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                isRowGroup: boolean;
                index: number;
                level: number;
                cancel: boolean;
            }


            export interface IRangeSortedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                col: number;
                ascending: boolean;
            }


            export interface IRangeSortingEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                col: number;
                ascending: boolean;
            }


            export interface IRowChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                row: number;
                sheetArea: SheetArea;
                propertyName: string;
                oldValue: any;
                newValue: any;
                count?: number;
            }


            export interface IRowHeightChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                rowList: any[];
                header: boolean;
            }


            export interface IRowHeightChangingEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                rowList: any[];
                header: boolean;
                cancel: boolean;
            }


            export interface ISelectionChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                oldSelections: Sheets.Range[];
                newSelections: Sheets.Range[];
            }


            export interface ISelectionChangingEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                oldSelections: any[];
                newSelections: any[];
            }


            export interface ISetBorderOptions{
                all?: boolean;
                left?: boolean;
                top?: boolean;
                right?: boolean;
                bottom?: boolean;
                outline?: boolean;
                inside?: boolean;
                innerHorizontal?: boolean;
                innerVertical?: boolean
            }


            export interface ISheetDefaultOption{
                rowHeight?: number;
                colHeaderRowHeight?: number;
                colWidth?: number;
                rowHeaderColWidth?: number;
            }


            export interface ISheetNameChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                oldValue: string;
                newValue: string;
            }


            export interface ISheetNameChangingEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                oldValue: string;
                newValue: string;
                cancel: boolean;
            }


            export interface ISheetTabClickEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                sheetTabIndex: number;
            }


            export interface ISheetTabDoubleClickEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                sheetTabIndex: number;
            }


            export interface ISlicerChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                slicer: GC.Spread.Sheets.Slicers.Slicer;
                propertyName: string;
            }


            export interface ISparklineChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                sparkline: Sparklines.Sparkline;
            }


            export interface ITableFilteredEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                table: GC.Spread.Sheets.Tables.Table;
                tableCol: number;
                filterValues: any[];
            }


            export interface ITableFilteringEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                table: GC.Spread.Sheets.Tables.Table;
                tableCol: number;
                filterValues: any[];
            }


            export interface ITopRowChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                oldTopRow: number;
                newTopRow: number;
            }


            export interface ITouchToolStripOpeningEventArgs{
                x: number;
                y: number;
                handled: boolean;
            }


            export interface IUserFormulaEnteredEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                row: number;
                col: number;
                formula: string;
            }


            export interface IUserZoomingEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                oldZoomFactor: number;
                newZoomFactor: number;
            }


            export interface IValidationErrorEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                row: number;
                col: number;
                validator: DataValidation.DefaultDataValidator;
                validationResult: DataValidation.DataValidationResult;
            }


            export interface IValueChangedEventArgs{
                sheet: GC.Spread.Sheets.Worksheet;
                sheetName: string;
                row: number;
                col: number;
                oldValue: any;
                newValue: any;
            }


            export interface IWorkbookOptions{
                allowUserDragDrop :boolean ; // Whether to allow the user to drag and drop range data.
                allowUserDragFill :boolean ; // Whether to allow the user to drag fill a range.
                allowUserZoom :boolean ; // Whether to zoom the display by scrolling the mouse wheel while pressing the Ctrl key.
                allowUserResize :boolean ; // Whether to allow the user to resize columns and rows.
                allowUndo :boolean ; // Whether to allow the user to undo edits.
                allowSheetReorder :boolean ; // Whether the user can reorder the sheets in the Spread component.
                defaultDragFillType :GC.Spread.Sheets.Fill.AutoFillType; // The default fill type.
                showDragFillSmartTag :boolean ; // Whether to display the drag fill dialog.
                showHorizontalScrollbar :boolean ; // Whether to display the horizontal scroll bar.
                showVerticalScrollbar :boolean ; // Whether to display the vertical scroll bar.
                scrollbarShowMax :boolean ; // Whether the displayed scroll bars are based on the entire number of columns and rows in the sheet.
                scrollbarMaxAlign :boolean ; // Whether the scroll bar aligns with the last row and column of the active sheet.
                tabStripVisible :boolean ; // Whether to display the sheet tab strip.
                tabStripRatio :number; // The width of the tab strip expressed as a percentage of the overall horizontal scroll bar width.
                tabEditable :boolean ; // Whether to allow the user to edit the sheet tab strip.
                newTabVisible :boolean ; // Whether the spreadsheet displays the special tab to let users insert new sheets.
                tabNavigationVisible :boolean ; // Whether to display the sheet tab navigation.
                cutCopyIndicatorVisible :boolean ; // Whether to display an indicator when copying or cutting the selected item.
                cutCopyIndicatorBorderColor :string; // The border color for the indicator displayed when the user cuts or copies the selection.
                backColor :string; // A color string used to represent the background color of the Spread component, such as "red", "#FFFF00", "rgb(255,0,0)", "Accent 5", and so on.
                backgroundImage :string; // The background image of the Spread component.
                backgroundImageLayout :GC.Spread.Sheets.ImageLayout; // The background image layout for the Spread component.
                grayAreaBackColor :string; // A color string used to represent the background color of the gray area , such as "red", "#FFFF00", "rgb(255,0,0)", "Accent 5", and so on.
                showResizeTip :GC.Spread.Sheets.ShowResizeTip; // How to display the resize tip.
                showDragDropTip :boolean ; // Whether to display the drag-drop tip.
                showDragFillTip :boolean ; // Whether to display the drag-fill tip.
                showScrollTip :GC.Spread.Sheets.ShowScrollTip; // How to display the scroll tip.
                scrollIgnoreHidden :boolean ; // Whether the scroll bar ignores hidden rows or columns.
                highlightInvalidData :boolean ; // Whether to highlight invalid data.
                useTouchLayout :boolean ; // Whether to use touch layout to present the Spread component.
                hideSelection :boolean ; // Whether to display the selection highlighting when the Spread component does not have focus.
                resizeZeroIndicator :GC.Spread.Sheets.ResizeZeroIndicator; // The drawing policy when the row or column is resized to zero.
                allowUserEditFormula :boolean ; // Whether the user can edit formulas in a cell in the spreadsheet.
                enableFormulaTextbox :boolean ; // Whether to enable the formula text box in the spreadsheet.
                autoFitType :GC.Spread.Sheets.AutoFitType; // Whether content will be formatted to fit in cells or in cells and headers.
                referenceStyle :GC.Spread.Sheets.ReferenceStyle; // the style for cell and range references in cell formulas on this sheet.
            }


            export interface IWorkSheetGridlineOption{
                color: string; //The grid line color
                showVerticalGridline: boolean; //Whether to show the vertical grid line.
                showHorizontalGridline: boolean; //Whether to show the horizontal grid line.
            }


            export interface IWorksheetOptions{
                allowCellOverflow: boolean; //indicates whether data can overflow into adjacent empty cells.
                sheetTabColor: string; //A color string used to represent the sheet tab color, such as "red", "#FFFF00", "rgb(255,0,0)", "Accent 5", and so on.
                frozenlineColor: string; //A color string used to represent the frozen line color, such as "red", "#FFFF00", "rgb(255,0,0)", "Accent 5", and so on.
                clipBoardOptions: GC.Spread.Sheets.ClipboardPasteOptions; //The clipboard option.
                gridline: GC.Spread.Sheets.IWorkSheetGridlineOption; //The grid line's options.
                rowHeaderVisible: boolean; //Indicates whether the row header is visible.
                colHeaderVisible: boolean; //Indicates whether the column header is visible.
                rowHeaderAutoText: GC.Spread.Sheets.HeaderAutoText; //Indicates whether the row header displays letters or numbers or is blank.
                colHeaderAutoText: GC.Spread.Sheets.HeaderAutoText; //Indicates whether the column header displays letters or numbers or is blank.
                rowHeaderAutoTextIndex: GC.Spread.Sheets.HeaderAutoText; //Specifies which row header column displays the automatic text when there are multiple row header columns.
                colHeaderAutoTextIndex: GC.Spread.Sheets.HeaderAutoText; //Specifies which column header row displays the automatic text when there are multiple column header rows.
                isProtected: boolean; //Indicates whether cells on this sheet that are marked as protected cannot be edited.
                protectionOptions: GC.Spread.Sheets.IProtectionOptions; //A value that indicates the elements that you want users to be able to change.
                selectionBackColor: string; //The selection's background color for the sheet.
                selectionBorderColor: string; //The selection's border color for the sheet.
            }

            /**
             * Represents whether the component automatically resizes cells or headers.
             * @enum {number}
             */
            export enum AutoFitType{
                /**
                 *  The component autofits cells.
                 */
                cell= 0,
                /**
                 *   The component autofits cells and headers.
                 */
                cellWithHeader= 1
            }

            /**
             * Specifies what data is pasted from the Clipboard.
             * @enum {number}
             */
            export enum ClipboardPasteOptions{
                /**
                 * Pastes all data objects, including values, formatting, and formulas.
                 */
                all= 0,
                /**
                 * Pastes only values.
                 */
                values= 1,
                /**
                 * Pastes only formatting.
                 */
                formatting= 2,
                /**
                 * Pastes only formulas.
                 */
                formulas= 3
            }

            /**
             * Specifies which headers are included when data is copied to or pasted.
             * @enum {number}
             */
            export enum CopyPasteHeaderOptions{
                /**
                 * Includes neither column nor row headers when data is copied; does not overwrite selected column or row headers when data is pasted.
                 */
                noHeaders= 0,
                /**
                 * Includes selected row headers when data is copied; overwrites selected row headers when data is pasted.
                 */
                rowHeaders= 1,
                /**
                 * Includes selected column headers when data is copied; overwrites selected column headers when data is pasted.
                 */
                columnHeaders= 2,
                /**
                 * Includes selected headers when data is copied; overwrites selected headers when data is pasted.
                 */
                allHeaders= 3
            }

            /**
             * Specifies the copy to option.
             * @enum {number}
             */
            export enum CopyToOptions{
                /**
                 * Indicates the type of data is pure data.
                 */
                value= 0x01,
                /**
                 * Indicates the type of data is a formula.
                 */
                formula= 0x02,
                /**
                 * Indicates the type of data is a comment.
                 */
                comment= 0x04,
                /**
                 * Indicates to copy a range group.
                 */
                outline= 0x08,
                /**
                 * Indicates the type of data is a sparkline.
                 */
                sparkline= 0x10,
                /**
                 * Indicates to copy a span.
                 */
                span= 0x20,
                /**
                 * Indicates the type of data is a style.
                 */
                style= 0x40,
                /**
                 * Indicates the type of data is a tag.
                 */
                tag= 0x80,
                /**
                 * Indicates the type of data is a binding path.
                 */
                bindingPath= 0x100,
                /**
                 * Indicates the type of data is a conditional format.
                 */
                conditionalFormat= 0x200,
                /**
                 * Indicates all types of data.
                 */
                all= 0x3ff
            }

            /**
             * Specifies the editor status.
             * @enum {number}
             */
            export enum EditorStatus{
                /**
                 * Cell is in Ready mode.
                 */
                ready= 0,
                /**
                 * Cell is in editing mode and can commit the input value and navigate to or select other cells when invoking navigation or selection actions.
                 */
                enter= 1,
                /**
                 * Cell is in editing mode and cannot commit the input value and navigate to or select other cells.
                 */
                edit= 2
            }

            /**
             * Specifies which default labels are displayed in headers.
             * @enum {number}
             */
            export enum HeaderAutoText{
                /**
                 *  Displays blanks in the headers.
                 */
                blank= 0,
                /**
                 *  Displays numbers in the headers.
                 */
                numbers= 1,
                /**
                 *  Displays letters in the headers.
                 */
                letters= 2
            }

            /**
             * Specifies the horizontal alignment.
             * @enum {number}
             */
            export enum HorizontalAlign{
                /**
                 *  Indicates that the cell content is left-aligned.
                 */
                left= 0,
                /**
                 *  Indicates that the cell content is centered.
                 */
                center= 1,
                /**
                 *  Indicates that the cell content is right-aligned.
                 */
                right= 2,
                /**
                 *  Indicates that the horizontal alignment is based on the value type.
                 */
                general= 3
            }

            /**
             * Specifies the horizontal position of the cell or column in the component.
             * @enum {number}
             */
            export enum HorizontalPosition{
                /**
                 *  Positions the cell or column to the left.
                 */
                left= 0,
                /**
                 *  Positions the cell or column in the center.
                 */
                center= 1,
                /**
                 *  Positions the cell or column to the right.
                 */
                right= 2,
                /**
                 *  Positions the cell or column to the nearest edge.
                 */
                nearest= 3
            }

            /**
             * Defines the background image layout.
             * @enum {number}
             */
            export enum ImageLayout{
                /** Specifies that the background image fills the area.
                 * @type {number}
                 */
                stretch= 0,
                /** Specifies that the background image displays in the center of the area.
                 * @type {number}
                 */
                center= 1,
                /** Specifies that the background image displays in the area with its original aspect ratio.
                 * @type {number}
                 */
                zoom= 2,
                /** Specifies that the background image displays in the upper left corner of the area with its original size.
                 * @type {number}
                 */
                none= 3
            }

            /**
             * Defines the IME mode to control the state of the Input Method Editor (IME).
             * @enum {number}
             */
            export enum ImeMode{
                /**
                 * No change is made to the current input method editor state.
                 */
                auto= 0x01,
                /** All characters are entered through the IME. Users can still deactivate the IME.
                 */
                active= 0x02,
                /**
                 * All characters are entered without IME. Users can still activate the IME.
                 */
                inactive= 0x04,
                /**
                 * The input method editor is disabled and may not be activated by the user.
                 */
                disabled= 0x00
            }

            /**
             * Identifies which operation was invalid.
             * @enum {number}
             */
            export enum InvalidOperationType{
                /**
                 * Specifes the formula is invalid.
                 */
                setFormula= 0,
                /**
                 * Specifies the copy paste is invalid.
                 */
                copyPaste= 1,
                /**
                 * Specifies the drag fill is invalid.
                 */
                dragFill= 2,
                /**
                 * Specifies the drag drop is invalid.
                 */
                dragDrop= 3,
                /**
                 * Specifies the insert row is invalid.
                 */
                changePartOfArrayFormula= 4,
                /**
                 * Specifies the changed sheet name is invalid.
                 */
                changeSheetName= 5
            }

            /**
             * Specifies the cell label position.
             * @enum {number}
             */
            export enum LabelAlignment{
                /**
                 *  Indicates that the cell label position is top-left.
                 */
                topLeft= 0,
                /**
                 *  Indicates that the cell label position is top-center.
                 */
                topCenter= 1,
                /**
                 *  Indicates that the cell label position is top-right.
                 */
                topRight= 2,
                /**
                 *  Indicates that the cell label position is bottom-left.
                 */
                bottomLeft= 3,
                /**
                 *  Indicates that the cell label position is bottom-center.
                 */
                bottomCenter= 4,
                /**
                 *  Indicates that the cell label position is bottom-right.
                 */
                bottomRight= 5
            }

            /**
             * Specifies the cell label visibility.
             * @enum {number}
             */
            export enum LabelVisibility{
                /**
                 *  Specifies to always show the watermark in the padding area and not to show the watermark in the cell area, regardless of the cell value.
                 */
                visible= 0,
                /**
                 *  Specifies to not show the watermark in the padding area, but to show the watermark in the cell area based on a value condition.
                 */
                hidden= 1,
                /**
                 *  Specifies to show the watermark in the padding area when the cell has a value or to show the watermark in the cell area if the cell does not have a value.
                 */
                auto= 2
            }

            /**
             * Specifies the line drawing style for the border.
             * @enum {number}
             */
            export enum LineStyle{
                /**
                 * Indicates a border line without a style.
                 */
                empty= 0,
                /**
                 *  Indicates a border line with a solid thin line.
                 */
                thin= 1,
                /**
                 *  Indicates a medium border line with a solid line.
                 */
                medium= 2,
                /**
                 *  Indicates a border line with dashes.
                 */
                dashed= 3,
                /**
                 *  Indicates a border line with dots.
                 */
                dotted= 4,
                /**
                 *  Indicates a thick border line with a solid line.
                 */
                thick= 5,
                /**
                 *  Indicates a double border line.
                 */
                double= 6,
                /**
                 *  Indicates a border line with all dots.
                 */
                hair= 7,
                /**
                 *  Indicates a medium border line with dashes.
                 */
                mediumDashed= 8,
                /**
                 *  Indicates a border line with dash-dot.
                 */
                dashDot= 9,
                /**
                 *  Indicates a medium border line with dash-dot-dot.
                 */
                mediumDashDot= 10,
                /**
                 *  Indicates a border line with dash-dot-dot.
                 */
                dashDotDot= 11,
                /**
                 *  Indicates a medium border line with dash-dot-dot.
                 */
                mediumDashDotDot= 12,
                /**
                 *  Indicates a slanted border line with dash-dot.
                 */
                slantedDashDot= 13
            }

            /**
             * Defines the type of action that raised the RangeChanged event.
             * @enum {number}
             */
            export enum RangeChangedAction{
                /**
                 * Indicates drag drop undo action.
                 */
                dragDrop= 0,
                /**
                 * Indicates drag fill undo action.
                 */
                dragFill= 1,
                /**
                 * Indicates clear range value undo action.
                 */
                clear= 2,
                /**
                 * Indicates paste undo action.
                 */
                paste= 3,
                /**
                 * Indicates sorting a range of cells.
                 */
                sort= 4,
                /**
                 * Indicates setting a formula in a specified range of cells .
                 */
                setArrayFormula= 5,
                /**
                 * Indicates setting a formula in a specified range of cells .
                 */
                evaluateFormula= 6
            }

            /**
             * Specifies the formula reference style.
             * @enum {number}
             */
            export enum ReferenceStyle{
                /**
                 * Indicates a1 style.
                 */
                a1= 0,
                /**
                 * Indicates r1c1 style.
                 */
                r1c1= 1
            }

            /**
             * Specifies the drawing policy of the row or column when it is resized to zero.
             * @enum {number}
             */
            export enum ResizeZeroIndicator{
                /**
                 *  Uses the current drawing policy when the row or column is resized to zero.
                 */
                default= 0,
                /**
                 * Draws two short lines when the row or column is resized to zero.
                 */
                enhanced= 1
            }

            /**
             * Specifies how users can select items in the control.
             * @enum {number}
             */
            export enum SelectionPolicy{
                /**
                 * Allows users to only select single items.
                 */
                single= 0,
                /**
                 * Allows users to select single items and ranges of items, but not multiple ranges.
                 */
                range= 1,
                /**
                 * Allows users to select single items and ranges of items, including multiple ranges.
                 */
                multiRange= 2
            }

            /**
             * Specifies the smallest unit users or the application can select.
             * @enum {number}
             */
            export enum SelectionUnit{
                /**
                 * Indicates that the smallest unit that can be selected is a cell.
                 */
                cell= 0,
                /**
                 * Indicates that the smallest unit that can be selected is a row.
                 */
                row= 1,
                /**
                 * Indicates that the smallest unit that can be selected is a column.
                 */
                column= 2
            }

            /**
             * Specifies the sheet area.
             * @enum {number}
             */
            export enum SheetArea{
                /**
                 * Indicates the sheet corner.
                 */
                corner= 0,
                /**
                 * Indicates the column header.
                 */
                colHeader= 1,
                /**
                 * Indicates the row header.
                 */
                rowHeader= 2,
                /**
                 * Indicates the viewport.
                 */
                viewport= 3
            }

            /**
             * Defines how the resize tip is displayed.
             * @enum {number}
             */
            export enum ShowResizeTip{
                /** Specifies that no resize tip is displayed.
                 * @type {number}
                 */
                none= 0,
                /** Specifies that only the horizontal resize tip is displayed.
                 * @type {number}
                 */
                column= 1,
                /** Specifies that only the vertical resize tip is displayed.
                 * @type {number}
                 */
                row= 2,
                /** Specifies that horizontal and vertical resize tips are displayed.
                 * @type {number}
                 */
                both= 3
            }

            /**
             * Specifies how the scroll tip is displayed.
             * @enum {number}
             */
            export enum ShowScrollTip{
                /** Specifies that no scroll tip is displayed.
                 * @type {number}
                 */
                none= 0,
                /** Specifies that only the horizontal scroll tip is displayed.
                 * @type {number}
                 */
                horizontal= 1,
                /** Specifies that only the vertical scroll tip is displayed.
                 * @type {number}
                 */
                vertical= 2,
                /** Specifies that horizontal and vertical scroll tips are displayed.
                 * @type {number}
                 */
                both= 3
            }

            /**
             * Specifies the type of sorting to perform.
             * @enum {number}
             */
            export enum SortState{
                /** Indicates the sorting is disabled.
                 * @type {number}
                 */
                none= 0,
                /** Indicates the sorting is ascending.
                 * @type {number}
                 */
                ascending= 1,
                /** Indicates the sorting is descending.
                 * @type {number}
                 */
                descending= 2
            }

            /**
             * Represents the storage data type.
             * @enum {number}
             */
            export enum StorageType{
                /**
                 *  Indicates the storage data type is pure value.
                 */
                data= 0x01,
                /**
                 *  Indicates the storage data type is style.
                 */
                style= 0x02,
                /**
                 *  Indicates the storage data type is comment.
                 */
                comment= 0x04,
                /**
                 *  Indicates the storage data type is tag.
                 */
                tag= 0x08,
                /**
                 *  Indicates the storage data type is sparkline.
                 */
                sparkline= 0x10,
                /**
                 *  Indicates the storage data type is the axis information.
                 */
                axis= 0x20,
                /**
                 *  Indicates the storage data type is data binding path.
                 */
                bindingPath= 0x40
            }

            /**
             * Defines the type of the text decoration.
             * @enum {number}
             */
            export enum TextDecorationType{
                /** Specifies to display a line below the text.
                 */
                underline= 1,
                /** Specifies to display a line through the text.
                 */
                lineThrough= 2,
                /** Specifies to display a line above the text.
                 */
                overline= 4,
                /** Specifies normal text.
                 */
                none= 0
            }

            /**
             * Specifies the vertical alignment.
             * @enum {number}
             */
            export enum VerticalAlign{
                /**
                 *  Indicates that the cell content is top-aligned.
                 */
                top= 0,
                /**
                 *  Indicates that the cell content is centered.
                 */
                center= 1,
                /**
                 *  Indicates that the cell content is bottom-aligned.
                 */
                bottom= 2
            }

            /**
             * Specifies the vertical position of the cell or row in the component.
             * @enum {number}
             */
            export enum VerticalPosition{
                /**
                 *  Positions the cell or row at the top.
                 */
                top= 0,
                /**
                 *  Positions the cell or row in the center.
                 */
                center= 1,
                /**
                 *  Positions the cell or row at the bottom.
                 */
                bottom= 2,
                /**
                 *  Positions the cell or row at the nearest edge.
                 */
                nearest= 3
            }

            /**
             * Specifies the visual state.
             * @enum {number}
             */
            export enum VisualState{
                /**
                 * Indicates normal visual state.
                 */
                normal= 0,
                /**
                 * Indicates highlight visual state.
                 */
                highlight= 1,
                /**
                 * Indicates selected visual state.
                 */
                selected= 2,
                /**
                 * Indicates active visual state.
                 */
                active= 3,
                /**
                 * Indicates hover visual state.
                 */
                hover= 4
            }


            export class CellRange{
                /**
                 * Represents a cell range in a sheet.
                 * @class
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that contains this cell range.
                 * @param {number} row The row index of the cell.
                 * @param {number} col The column index of the cell.
                 * @param {number} rowCount The row count of the cell. If you do not provide this parameter, it defaults to <b>1</b>.
                 * @param {number} colCount The column count of the cell. If you do not provide this parameter, it defaults to <b>1</b>.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If you do not provide this parameter, it defaults to <b>viewport</b>.
                 * If row is -1 and rowCount is -1, the range represents columns. For example, new GC.Spread.Sheets.CellRange(-1,4,-1,6) represents columns "E:J".
                 * If col is -1 and colCount is -1, the range represents rows. For example, new GC.Spread.Sheets.CellRange(4,-1,6,-1) represents rows "5:10".
                 */
                constructor(sheet:  GC.Spread.Sheets.Worksheet,  row:  number,  col:  number,  rowCount?:  number,  colCount?:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea);
                /**
                 * Gets the starting column index.
                 * @type {number}
                 */
                col: number;
                /**
                 * Gets the column count.
                 * @type {number}
                 */
                colCount: number;
                /**
                 * Gets the starting row index.
                 *@type {number}
                 */
                row: number;
                /**
                 * Gets the row count.
                 * @type {number}
                 */
                rowCount: number;
                /**
                 * Gets the sheet that contains this cell range.
                 * @type {GC.Spread.Sheets.Worksheet}
                 */
                sheet: GC.Spread.Sheets.Worksheet;
                /**
                 * Gets the area that contains this cell range.
                 * @type {GC.Spread.Sheets.SheetArea}
                 */
                sheetArea: GC.Spread.Sheets.SheetArea;
                /**
                 * Gets or sets the background color for the cell, such as "red", "#FFFF00", "rgb(255,0,0)", "Accent 5", and so on.
                 * @param {string} value The cell background color.
                 * @returns {string|GC.Spread.Sheets.CellRange} If no value is set, returns the cell background color; otherwise, returns the cell.
                 */
                backColor(value?:  string): any;
                /**
                 * Gets or sets the background image for the cell.
                 * @param {string} value The cell background image.
                 * @returns {string|GC.Spread.Sheets.CellRange} If no value is set, returns the cell background image; otherwise, returns the cell.
                 */
                backgroundImage(value?:  string): any;
                /**
                 * Gets or sets the background image layout for the cell.
                 * @param {GC.Spread.Sheets.ImageLayout} value The cell background image layout.
                 * @returns {GC.Spread.Sheets.ImageLayout|GC.Spread.Sheets.CellRange} If no value is set, returns the cell background image layout; otherwise, returns the cell.
                 */
                backgroundImageLayout(value?:  GC.Spread.Sheets.ImageLayout): any;
                /**
                 * Gets or sets the binding path for cell binding.
                 * @param {string} path The binding path for cell binding.
                 * @returns {string | GC.Spread.Sheets.Worksheet} If no value is set, returns the binding path for cell binding; otherwise, returns the worksheet.
                 */
                bindingPath(path?:  string): any;
                /**
                 * Gets or sets the bottom border of the cell.
                 * @param {GC.Spread.Sheets.LineBorder} value The cell bottom border line.
                 * @returns {GC.Spread.Sheets.LineBorder|GC.Spread.Sheets.CellRange} If no value is set, returns the cell bottom border line; otherwise, returns the cell.
                 */
                borderBottom(value?:  GC.Spread.Sheets.LineBorder): any;
                /**
                 * Gets or sets the left border of the cell.
                 * @param {GC.Spread.Sheets.LineBorder} value The cell left border line.
                 * @returns {GC.Spread.Sheets.LineBorder|GC.Spread.Sheets.CellRange} If no value is set, returns the cell left border line; otherwise, returns the cell.
                 */
                borderLeft(value?:  GC.Spread.Sheets.LineBorder): any;
                /**
                 * Gets or sets the right border of the cell.
                 * @param {GC.Spread.Sheets.LineBorder} value The cell right border line.
                 * @returns {GC.Spread.Sheets.LineBorder|GC.Spread.Sheets.CellRange} If no value is set, returns the cell right border line; otherwise, returns the cell.
                 */
                borderRight(value?:  GC.Spread.Sheets.LineBorder): any;
                /**
                 * Gets or sets the top border of the cell.
                 * @param {GC.Spread.Sheets.LineBorder} value The cell top border line.
                 * @returns {GC.Spread.Sheets.LineBorder|GC.Spread.Sheets.CellRange} If no value is set, returns the cell top border line; otherwise, returns the cell.
                 */
                borderTop(value?:  GC.Spread.Sheets.LineBorder): any;
                /**
                 * Gets or sets the cell padding.
                 * @param {string} value The cell padding.
                 * @returns {string|GC.Spread.Sheets.CellRange} If no value is set, returns the value of the cell padding; otherwise, returns the cell.
                 */
                cellPadding(value?:  string): any;
                /**
                 * Gets or sets the cell type of the cell.
                 * @param {GC.Spread.Sheets.CellTypes.Base} value The cell type.
                 * @returns {GC.Spread.Sheets.CellTypes.Base|GC.Spread.Sheets.CellRange} If no value is set, returns the cell type; otherwise, returns the cell.
                 */
                cellType(value?:  GC.Spread.Sheets.CellTypes.Base): any;
                /**
                 * Clears the specified area.
                 * @param {GC.Spread.Sheets.StorageType} type The clear type.
                 */
                clear(type:  GC.Spread.Sheets.StorageType): void;
                /**
                 * Gets or sets the comment for the cell.
                 * @param {GC.Spread.Sheets.Comments.Comment} value The comment to set in the cell.
                 * @returns {GC.Spread.Sheets.Comments.Comment} If no value is set, returns the comment in the cell; otherwise, returns the cell range.
                 */
                comment(value?:  GC.Spread.Sheets.Comments.Comment): any;
                /**
                 * Gets or sets the font for the cell, such as "normal normal normal 20px/normal Arial".
                 * @param {string} value The cell font.
                 * @returns {string|GC.Spread.Sheets.CellRange} If no value is set, returns the cell's font; otherwise, returns the cell.
                 */
                font(value?:  string): any;
                /**
                 * Gets or sets the color of the text in the cell, such as "red", "#FFFF00", "rgb(255,0,0)", "Accent 5", and so on.
                 * @param {string} value The color of the text.
                 * @returns {string|GC.Spread.Sheets.CellRange} If no value is set, returns the cell foreground color; otherwise, returns the cell.
                 */
                foreColor(value?:  string): any;
                /**
                 * Gets or sets the formatter for the cell.
                 * @param {string | GC.Spread.Formatter.GeneralFormatter} value The cell formatter string or object.
                 * @returns {string | GC.Spread.Formatter.GeneralFormatter |GC.Spread.Sheets.CellRange} If no value is set, returns the cell formatter string or object; otherwise, returns the cell.
                 */
                formatter(value?:  any): any;
                /**
                 * Gets or sets the formula for the cell.
                 * @param {string} value The cell formula.
                 * @returns {string|GC.Spread.Sheets.CellRange} If no value is set, returns the cell formula; otherwise, returns the cell.
                 */
                formula(value?:  string): any;
                /**
                 * Gets or sets the horizontal alignment of the contents of the cell.
                 * @param {GC.Spread.Sheets.HorizontalAlign} value The horizontal alignment.
                 * @returns {GC.Spread.Sheets.HorizontalAlign|GC.Spread.Sheets.CellRange} If no value is set, returns the horizontal alignment of the contents of the cell; otherwise, returns the cell.
                 */
                hAlign(value?:  GC.Spread.Sheets.HorizontalAlign): any;
                /**
                 * Gets or sets the height of the row in pixels.
                 * @param {number} value The cell row height.
                 * @returns {number|GC.Spread.Sheets.CellRange} If no value is set, returns the row height; otherwise, returns the row.
                 */
                height(value?:  number): any;
                /**
                 * Gets or sets the imeMode of the cell.
                 * @param {GC.Spread.Sheets.ImeMode} value The cell imeMode.
                 * @returns {GC.Spread.Sheets.ImeMode|GC.Spread.Sheets.CellRange} If no value is set, returns the cell imeMode; otherwise, returns the cell.
                 */
                imeMode(value?:  GC.Spread.Sheets.ImeMode): any;
                /**
                 * Gets or sets the cell label options.
                 * @param {Object} value The cell label options.<br />
                 ** alignment {GC.Spread.Sheets.LabelAlignment} The cell label position.<br />
                 ** visibility {GC.Spread.Sheets.LabelVisibility} The cell label visibility.<br />
                 ** font {string} The cell label font.<br />
                 ** foreColor {string} The cell label forecolor.<br />
                 ** margin {string} The cell label margin.
                 * @returns {Object|GC.Spread.Sheets.CellRange} If no value is set, returns the value of the cell label options; otherwise, returns the cell.
                 */
                labelOptions(value?:  Object): any;
                /**
                 * Gets or sets whether the cell is locked. When the sheet is protected, the locked cell cannot be edited.
                 * @param {boolean} value Set to <c>true</c> to lock the cell.
                 * @returns {boolean|GC.Spread.Sheets.CellRange} If no value is set, returns whether the cell is locked; otherwise, returns the cell.
                 */
                locked(value?:  boolean): any;
                /**
                 * Gets or sets whether the row or column can be resized by the user.
                 * @param {boolean} value Set to <c>true</c> to make the row resizable.
                 * @returns {boolean|GC.Spread.Sheets.CellRange} If no value is set, returns whether the user can resize the row; otherwise, returns the row or column.
                 */
                resizable(value?:  boolean): any;
                /**
                 * Sets the border for the specified area.
                 * @param {GC.Spread.Sheets.LineBorder} border The border line.
                 * @param {Object} option Determines which part of the cell range to set, the option object contains {all:true, left:true, top:true, right:true, bottom:true,outline:true,inside:true, innerHorizontal:true, innerVertical:true}
                 */
                setBorder(border:  GC.Spread.Sheets.LineBorder,  option:  GC.Spread.Sheets.ISetBorderOptions): void;
                /**
                 * Gets or sets whether the cell shrinks the text to fit the cell size.
                 * @param {boolean} value Set to <c>true</c> to have the cell shrink text to fit.
                 * @returns {boolean|GC.Spread.Sheets.CellRange} If no value is set, returns whether the cell shrinks the text to fit; otherwise, returns the cell.
                 */
                shrinkToFit(value?:  boolean): any;
                /**
                 * Gets or sets a value that indicates whether the user can set focus to the cell using the Tab key.
                 * @param {boolean} value Set to <c>true</c> to set focus to the cell using the Tab key.
                 * @returns {boolean|GC.Spread.Sheets.CellRange} If no value is set, returns whether the user can set focus to the cell using the Tab key; otherwise, returns the cell.
                 */
                tabStop(value?:  boolean): any;
                /**
                 * Gets or sets the tag for the cell.
                 * @param {Object} value The tag value.
                 * @returns {Object|GC.Spread.Sheets.CellRange} If no value is set, returns the tag value; otherwise, returns the cell.
                 */
                tag(value?:  any): any;
                /**
                 * Gets or sets the formatted text for the cell.
                 * @param {string} value The cell text.
                 * @returns {string|GC.Spread.Sheets.CellRange} If no value is set, returns the cell text; otherwise, returns the cell.
                 */
                text(value?:  string): any;
                /**
                 * Gets or sets the type of the decoration added to the cell's text.
                 * @param {GC.Spread.Sheets.TextDecorationType} value The type of the decoration.
                 * @returns {GC.Spread.Sheets.TextDecorationType|GC.Spread.Sheets.CellRange} If no value is set, returns the type of the decoration; otherwise, returns the cell.
                 */
                textDecoration(value?:  GC.Spread.Sheets.TextDecorationType): any;
                /**
                 * Gets or sets the text indent of the cell.
                 * @param {number}  value The cell text indent.
                 * @returns {number|GC.Spread.Sheets.CellRange} If no value is set, returns the cell text indent; otherwise, returns the cell.
                 */
                textIndent(value?:  number): any;
                /**
                 * Gets or sets the theme font for the cell.
                 * @param {string} value The cell's theme font.
                 * @returns {string|GC.Spread.Sheets.CellRange} If no value is set, returns the cell's theme font; otherwise, returns the cell.
                 */
                themeFont(value?:  string): any;
                /**
                 * Gets or sets the data validator for the cell.
                 * @param {GC.Spread.Sheets.DataValidation.DefaultDataValidator} value The cell data validator.
                 * @returns {GC.Spread.Sheets.DataValidation.DefaultDataValidator|GC.Spread.Sheets.CellRange} If no value is set, returns the cell data validator; otherwise, returns the cell.
                 */
                validator(value?:  GC.Spread.Sheets.DataValidation.DefaultDataValidator): any;
                /**
                 * Gets or sets the vertical alignment of the contents of the cell.
                 * @param {GC.Spread.Sheets.VerticalAlign} value The vertical alignment.
                 * @returns {GC.Spread.Sheets.VerticalAlign|GC.Spread.Sheets.CellRange} If no value is set, returns the vertical alignment of the contents of the cell; otherwise, returns the cell.
                 */
                vAlign(value?:  GC.Spread.Sheets.VerticalAlign): any;
                /**
                 * Gets or sets the unformatted value of the cell.
                 * @param {Object} value The cell value.
                 * @returns {Object|GC.Spread.Sheets.CellRange} If no value is set, returns the cell value; otherwise, returns the cell.
                 */
                value(value?:  any): any;
                /**
                 * Gets or sets whether the row or column is displayed.
                 * @param {boolean} value Set to <c>true</c> to make the row visible.
                 * @returns {boolean|GC.Spread.Sheets.CellRange} If no value is set, returns the visible of the row or column; otherwise, returns the row or column.
                 */
                visible(value?:  boolean): any;
                /**
                 * Gets or sets the content of the cell watermark.
                 * @param {string} value The content of the watermark.
                 * @returns {string|GC.Spread.Sheets.CellRange} If no value is set, returns the content of the watermark; otherwise, returns the cell.
                 */
                watermark(value?:  string): any;
                /**
                 * Gets or sets the width of the column in pixels.
                 * @param {number} value The column width.
                 * @returns {number|GC.Spread.Sheets.CellRange} If no value is set, returns the column width; otherwise, returns the column.
                 */
                width(value?:  number): any;
                /**
                 * Gets or sets whether the cell lets text wrap.
                 * @param {boolean} value Set to <c>true</c> to let text wrap within the cell.
                 * @returns {boolean|GC.Spread.Sheets.CellRange} If no value is set, returns whether the cell lets text wrap; otherwise, returns the cell.
                 */
                wordWrap(value?:  boolean): any;
            }

            export class ColorScheme{
                /**
                 * Creates a ColorScheme instance.
                 * @constructor
                 * @class
                 * @classdesc Represents the theme color.
                 * @param {string} name The owner that contains the named variable.
                 * @param {string} background1 The theme color for background1.
                 * @param {string} background2 The theme color for background2.
                 * @param {string} text1 The theme color for text1.
                 * @param {string} text2 The theme color for text2.
                 * @param {string} accent1 The theme color for accent1.
                 * @param {string} accent2 The theme color for accent2.
                 * @param {string} accent3 The theme color for accent3.
                 * @param {string} accent4 The theme color for accent4.
                 * @param {string} accent5 The theme color for accent5.
                 * @param {string} accent6 The theme color for accent6.
                 * @param {string} link The color of the link.
                 * @param {string} followedLink The color of the followedLink.
                 */
                constructor(name:  string,  background1:  string,  background2:  string,  text1:  string,  text2:  string,  accent1:  string,  accent2:  string,  accent3:  string,  accent4:  string,  accent5:  string,  accent6:  string,  link:  string,  followedLink:  string);
                /**
                 * Gets or sets the accent1 theme color of the color scheme.
                 * @param {string} value The accent1 theme color string.
                 * @returns {string|GC.Spread.Sheets.ColorScheme} If no value is set, returns the accent1 theme color; otherwise, returns the color scheme.
                 */
                accent1(value?:  string): any;
                /**
                 * Gets or sets the accent2 theme color of the color scheme.
                 * @param {string} value The accent2 theme color string.
                 * @returns {string|GC.Spread.Sheets.ColorScheme} If no value is set, returns the accent2 theme color; otherwise, returns the color scheme.
                 */
                accent2(value?:  string): any;
                /**
                 * Gets or sets the accent3 theme color of the color scheme.
                 * @param {string} value The accent3 theme color string.
                 * @returns {string|GC.Spread.Sheets.ColorScheme} If no value is set, returns the accent3 theme color; otherwise, returns the color scheme.
                 */
                accent3(value?:  string): any;
                /**
                 * Gets or sets the accent4 theme color of the color scheme.
                 * @param {string} value The accent4 theme color string.
                 * @returns {string|GC.Spread.Sheets.ColorScheme} If no value is set, returns the accent4 theme color; otherwise, returns the color scheme.
                 */
                accent4(value?:  string): any;
                /**
                 * Gets or sets the accent5 theme color of the color scheme.
                 * @param {string} value The accent5 theme color string.
                 * @returns {string|GC.Spread.Sheets.ColorScheme} If no value is set, returns the accent5 theme color; otherwise, returns the color scheme.
                 */
                accent5(value?:  string): any;
                /**
                 * Gets or sets the accent6 theme color of the color scheme.
                 * @param {string} value The accent6 theme color string.
                 * @returns {string|GC.Spread.Sheets.ColorScheme} If no value is set, returns the accent6 theme color; otherwise, returns the color scheme.
                 */
                accent6(value?:  string): any;
                /**
                 * Gets or sets the background1 theme color of the color scheme.
                 * @param {string} value The background1 theme color string.
                 * @returns {string|GC.Spread.Sheets.ColorScheme} If no value is set, returns the background1 theme color; otherwise, returns the color scheme.
                 */
                background1(value?:  string): any;
                /**
                 *  Gets or sets the background2 theme color of the color scheme.
                 * @param {string} value The background2 theme color string.
                 * @returns {string|GC.Spread.Sheets.ColorScheme} If no value is set, returns the background2 theme color; otherwise, returns the color scheme.
                 */
                background2(value?:  string): any;
                /**
                 * Gets or sets the followed hyperlink theme color of the color scheme.
                 * @param {string} value The followed hyperlink theme color string.
                 * @returns {string|GC.Spread.Sheets.ColorScheme} If no value is set, returns the followed hyperlink theme color; otherwise, returns the color scheme.
                 */
                followedHyperlink(value?:  string): any;
                /**
                 * Gets the color based on the theme color.
                 * @param {string} name The theme color name.
                 * @returns {string} The theme color.
                 */
                getColor(name: any): void;
                /**
                 * Gets or sets the hyperlink theme color of the color scheme.
                 * @param {string} value The hyperlink theme color string.
                 * @returns {string|GC.Spread.Sheets.ColorScheme} If no value is set, returns the hyperlink theme color; otherwise, returns the color scheme.
                 */
                hyperlink(value?:  string): any;
                /**
                 * Gets or sets the name of the color scheme.
                 * @param {string} value The name.
                 * @returns {string|GC.Spread.Sheets.ColorScheme} If no value is set, returns the name; otherwise, returns the color scheme.
                 */
                name(value?:  string): any;
                /**
                 * Gets or sets the textcolor1 theme color of the color scheme.
                 * @param {string} value The textcolor1 theme color string.
                 * @returns {string|GC.Spread.Sheets.ColorScheme} If no value is set, returns the textcolor1 theme color; otherwise, returns the color scheme.
                 */
                textColor1(value?:  string): any;
                /**
                 * Gets or sets the textcolor2 theme color of the color scheme.
                 * @param {string} value The textcolor2 theme color string.
                 * @returns {string|GC.Spread.Sheets.ColorScheme} If no value is set, returns the textcolor2 theme color; otherwise, returns the color scheme.
                 */
                textColor2(value?:  string): any;
            }

            export class Events{
                /**
                 * Defines the events supported in SpreadJS.
                 * @class
                 */
                constructor();
                /**
                 * Occurs when the user has changed the active sheet.
                 * @name GC.Spread.Sheets.Workbook#ActiveSheetChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} oldSheet The old sheet.
                 * @param {GC.Spread.Sheets.Worksheet} newSheet The new sheet.
                 */
                static ActiveSheetChanged: string;
                /**
                 * Occurs when the user is changing the active sheet.
                 * @name GC.Spread.Sheets.Workbook#ActiveSheetChanging
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} oldSheet The old sheet.
                 * @param {GC.Spread.Sheets.Worksheet} newSheet The new sheet.
                 * @param {boolean} cancel A value that indicates whether the operation should be canceled.
                 */
                static ActiveSheetChanging: string;
                /**
                 * Occurs when the user clicks a button, check box, or hyperlink in a cell.
                 * @name GC.Spread.Sheets.Workbook#ButtonClicked
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} row The row index of the cell.
                 * @param {number} col The column index of the cell.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area that contains the cell.
                 */
                static ButtonClicked: string;
                /**
                 * Occurs when a change is made to a cell in this sheet that may require the cell to be repainted.
                 * @name GC.Spread.Sheets.Worksheet#CellChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} row The row index of the cell.
                 * @param {number} col The column index of the cell.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheetArea of the cell.
                 * @param {string} propertyName The name of the cell's property that has changed.
                 */
                static CellChanged: string;
                /**
                 * Occurs when the user presses down the left mouse button in a cell.
                 * @name GC.Spread.Sheets.Worksheet#CellClick
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area the clicked cell is in.
                 * @param {number} row The row index of the clicked cell.
                 * @param {number} col The column index of the clicked cell.
                 */
                static CellClick: string;
                /**
                 * Occurs when the user presses down the left mouse button twice (double-clicks) in a cell.
                 * @name GC.Spread.Sheets.Worksheet#CellDoubleClick
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area the clicked cell is in.
                 * @param {number} row The row index of the clicked cell.
                 * @param {number} col The column index of the clicked cell.
                 */
                static CellDoubleClick: string;
                /**
                 * Occurs when a Clipboard change occurs that affects Spread.Sheets.
                 * @name GC.Spread.Sheets.Worksheet#ClipboardChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {Object} copyData The data from Spread.Sheets that has been set into the clipboard.
                 * @param {string} copyData.text The text string of the clipboard.
                 * @param {string} copyData.html The html string of the clipboard.
                 */
                static ClipboardChanged: string;
                /**
                 * Occurs when the Clipboard is changing due to a Spread.Sheets action.
                 * @name GC.Spread.Sheets.Worksheet#ClipboardChanging
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {Object} copyData The data from Spread.Sheets that is set into the clipboard.
                 * @param {string} copyData.text The text string of the clipboard.
                 * @param {string} copyData.html The html string of the clipboard.
                 * @param {boolean} cancel A value that indicates whether the operation should be canceled.
                 */
                static ClipboardChanging: string;
                /**
                 * Occurs when the user pastes from the Clipboard.
                 * @name GC.Spread.Sheets.Worksheet#ClipboardPasted
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.Range} cellRange The range that was pasted.
                 * @param {GC.Spread.Sheets.ClipboardPasteOptions} pasteOption The paste option that indicates what data is pasted from the clipboard: values, formatting, or formulas.
                 * @param {Object} pasteData The data from the clipboard that will be pasted into Spread.Sheets.
                 * @param {string} pasteData.text The text string of the clipboard.
                 * @param {string} pasteData.html The html string of the clipboard.
                 */
                static ClipboardPasted: string;
                /**
                 * Occurs when the user is pasting from the Clipboard.
                 * @name GC.Spread.Sheets.Worksheet#ClipboardPasting
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.Range} cellRange The range to paste.
                 * @param {GC.Spread.Sheets.ClipboardPasteOptions} pasteOption The paste option that indicates what data is pasted from the clipboard: values, formatting, or formulas.
                 * @param {Object} pasteData The data from the clipboard that will be pasted into Spread.Sheets.
                 * @param {string} pasteData.text The text string of the clipboard.
                 * @param {string} pasteData.html The html string of the clipboard.
                 * @param {boolean} cancel A value that indicates whether the operation should be canceled.
                 */
                static ClipboardPasting: string;
                /**
                 * Occurs when a change is made to a column or range of columns in this sheet that may require the column or range of columns to be repainted.
                 * @name GC.Spread.Sheets.Worksheet#ColumnChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheetArea of the column.
                 * @param {string} propertyName The name of the column's property that has changed.
                 */
                static ColumnChanged: string;
                /**
                 * Occurs when the column width has changed.
                 * @name GC.Spread.Sheets.Worksheet#ColumnWidthChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {Array} colList The list of columns whose widths have changed.
                 * @param {boolean} header Whether the columns are row header columns.
                 */
                static ColumnWidthChanged: string;
                /**
                 * Occurs when the column width is changing.
                 * @name GC.Spread.Sheets.Worksheet#ColumnWidthChanging
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {Array} colList The list of columns whose widths are changing.
                 * @param {boolean} header Whether the columns are row header columns.
                 * @param {boolean} cancel A value that indicates whether the operation should be canceled.
                 */
                static ColumnWidthChanging: string;
                /**
                 * Occurs when any comment has changed.
                 * @name GC.Spread.Sheets.Worksheet#CommentChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.Comments.Comment} comment The comment that triggered the event.
                 * @param {string} propertyName The name of the comment's property that has changed.
                 */
                static CommentChanged: string;
                /**
                 * Occurs when the user has removed the comment.
                 * @name GC.Spread.Sheets.Worksheet#CommentRemoved
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.Comments.Comment} comment The comment has been removed.
                 */
                static CommentRemoved: string;
                /**
                 * Occurs when the user is removing any comment.
                 * @name GC.Spread.Sheets.Worksheet#CommentRemoving
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.Comments.Comment} comment The comment is being removed.
                 * @param {boolean} cancel A value that indicates whether the operation should be canceled.
                 */
                static CommentRemoving: string;
                /**
                 * Occurs when the user is dragging and dropping a range of cells.
                 * @name GC.Spread.Sheets.Worksheet#DragDropBlock
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} fromRow The row index of the top left cell of the source range (range being dragged).
                 * @param {number} fromCol The column index of the top left cell of the source range (range being dragged).
                 * @param {number} toRow The row index of the top left cell of the destination range (where selection is dropped).
                 * @param {number} toCol The column index of the bottom right cell of the destination range (where selection is dropped).
                 * @param {number} rowCount The row count of the cell range being dragged.
                 * @param {number} colCount The column count of the cell range being dragged.
                 * @param {boolean} copy Whether the source range is copied.
                 * @param {boolean} insert Whether the source range is inserted.
                 * @param {GC.Spread.Sheets.CopyToOptions} copyOption The CopyOption value for the drag and drop operation.
                 * @param {boolean} cancel A value that indicates whether the operation should be canceled.
                 */
                static DragDropBlock: string;
                /**
                 * Occurs when the user completes dragging and dropping a range of cells.
                 * @name GC.Spread.Sheets.Worksheet#DragDropBlockCompleted
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} fromRow The row index of the top left cell of the source range (range being dragged).
                 * @param {number} fromCol The column index of the top left cell of the source range (range being dragged).
                 * @param {number} toRow The row index of the top left cell of the destination range (where selection is dropped).
                 * @param {number} toCol The column index of the bottom right cell of the destination range (where selection is dropped).
                 * @param {number} rowCount The row count of the cell range being dragged.
                 * @param {number} colCount The column count of the cell range being dragged.
                 * @param {boolean} copy Whether the source range is copied.
                 * @param {boolean} insert Whether the source range is inserted.
                 * @param {GC.Spread.Sheets.CopyToOptions} copyOption The CopyOption value for the drag and drop operation.
                 */
                static DragDropBlockCompleted: string;
                /**
                 * Occurs when the user is dragging to fill a range of cells.
                 * @name GC.Spread.Sheets.Worksheet#DragFillBlock
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.Range} fillRange The range used for the fill operation.
                 * @param {GC.Spread.Sheets.Fill.AutoFillType} autoFillType AutoFillType value used for the fill operation.
                 * @param {GC.Spread.Sheets.Fill.FillDirection} fillDirection The direction of the fill.
                 * @param {boolean} cancel A value that indicates whether the operation should be canceled.
                 */
                static DragFillBlock: string;
                /**
                 * Occurs when the user completes dragging to fill a range of cells.
                 * @name GC.Spread.Sheets.Worksheet#DragFillBlockCompleted
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.Range} fillRange The range used for the fill operation.
                 * @param {GC.Spread.Sheets.Fill.AutoFillType} autoFillType AutoFillType value used for the fill operation.
                 * @param {GC.Spread.Sheets.Fill.FillDirection} fillDirection The direction of the fill.
                 */
                static DragFillBlockCompleted: string;
                /**
                 * Occurs when a cell is in edit mode and the text is changed.
                 * @name GC.Spread.Sheets.Worksheet#EditChange
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} row The row index of cell.
                 * @param {number} col The column index of cell.
                 * @param {object} editingText The value from the current editor.
                 */
                static EditChange: string;
                /**
                 * Occurs when a cell leaves edit mode.
                 * @name GC.Spread.Sheets.Worksheet#EditEnded
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} row The row index of cell.
                 * @param {number} col The column index of cell.
                 * @param {object} editingText The value from the current editor.
                 */
                static EditEnded: string;
                /**
                 * Occurs when a cell is leaving edit mode.
                 * @name GC.Spread.Sheets.Worksheet#EditEnding
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} row The row index of cell.
                 * @param {number} col The column index of cell.
                 * @param {object} editor The current editor.
                 * @param {object} editingText The value from the current editor.
                 * @param {boolean} cancel A value that indicates whether the operation should be canceled.
                 */
                static EditEnding: string;
                /**
                 * Occurs when the editor's status has changed.
                 * @name GC.Spread.Sheets.Worksheet#EditorStatusChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.EditorStatus} oldStatus The old status of the editor.
                 * @param {GC.Spread.Sheets.EditorStatus} newStatus The new status of the editor.
                 */
                static EditorStatusChanged: string;
                /**
                 * Occurs when a cell is entering edit mode.
                 * @name GC.Spread.Sheets.Worksheet#EditStarting
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} row The row index of cell.
                 * @param {number} col The column index of cell.
                 * @param {boolean} cancel A value that indicates whether the operation should be canceled.
                 */
                static EditStarting: string;
                /**
                 * Occurs when the focus enters a cell.
                 * @name GC.Spread.Sheets.Worksheet#EnterCell
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} row The row index of the cell being entered.
                 * @param {number} col The column index of the cell being entered.
                 */
                static EnterCell: string;
                /**
                 * Occurs when any floating object has changed.
                 * @name GC.Spread.Sheets.Worksheet#FloatingObjectsChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.FloatingObjects.FloatingObject} floatingObject The floating object that triggered the event.
                 * @param {string} propertyName The name of the floating object's property that has changed.
                 */
                static FloatingObjectChanged: string;
                /**
                 * Occurs when the custom floating object content is loaded.
                 * @name GC.Spread.Sheets.Worksheet#FloatingObjectLoaded
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.FloatingObjects.FloatingObject} floatingObject The custom floating object that triggered the event.
                 * @param {HTMLElement} element The HTMLElement of the custom floating object.
                 */
                static FloatingObjectLoaded: string;
                /**
                 * Occurs when the user has removed the floating object.
                 * @name GC.Spread.Sheets.Worksheet#FloatingObjectRemoved
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.FloatingObjects.FloatingObject} floatingObject The floating object has been removed.
                 */
                static FloatingObjectRemoved: string;
                /**
                 * Occurs when the user is removing any floating object.
                 * @name GC.Spread.Sheets.Worksheet#FloatingObjectRemoving
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.FloatingObjects.FloatingObject} floatingObject The floating object is being removed.
                 * @param {boolean} cancel A value that indicates whether the operation should be canceled.
                 */
                static FloatingObjectRemoving: string;
                /**
                 * Occurs when the selections of the floating object have changed.
                 * @name GC.Spread.Sheets.Worksheet#FloatingObjectsSelectionChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.FloatingObjects.FloatingObject} floatingObject The floating object that triggered the event.
                 */
                static FloatingObjectSelectionChanged: string;
                /**
                 * Occurs when an invalid operation is performed.
                 * @name GC.Spread.Sheets.Worksheet#InvalidOperation
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.InvalidOperationType} invalidType Indicates which operation was invalid.
                 * @param {string} message The description of the invalid operation.
                 */
                static InvalidOperation: string;
                /**
                 * Occurs when the focus leaves a cell.
                 * @name GC.Spread.Sheets.Worksheet#LeaveCell
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} row The row index of the cell being left.
                 * @param {number} col The column index of the cell being left.
                 * @param {boolean} cancel Whether the operation should be canceled.
                 */
                static LeaveCell: string;
                /**
                 * Occurs when the left column changes.
                 * @name GC.Spread.Sheets.Worksheet#LeftColumnChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} oldLeftCol The old left column index.
                 * @param {number} newLeftCol The new left column index.
                 */
                static LeftColumnChanged: string;
                /**
                 * Occurs when any picture has changed.
                 * @name GC.Spread.Sheets.Worksheet#PictureChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.FloatingObjects.Picture} picture The picture that triggered the event.
                 * @param {string} propertyName The name of the picture's property that has changed.
                 */
                static PictureChanged: string;
                /**
                 * Occurs when the selections of the picture have changed.
                 * @name GC.Spread.Sheets.Worksheet#PictureSelectionChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.FloatingObjects.Picture} picture The picture that triggered the event.
                 */
                static PictureSelectionChanged: string;
                /**
                 * Occurs when the cell range has changed.
                 * @name GC.Spread.Sheets.Worksheet#RangeChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} row The range's row index.
                 * @param {number} col The range's column index.
                 * @param {number} rowCount The range's row count.
                 * @param {number} colCount The range's column count.
                 * @param {Array.<Object>} changedCells The positions of the cells whose data has changed, each position has row and col.
                 * @param {GC.Spread.Sheets.RangeChangedAction} action The type of action that raises the RangeChanged event.
                 */
                static RangeChanged: string;
                /**
                 * Occurs when a column has just been automatically filtered.
                 * @name GC.Spread.Sheets.Worksheet#RangeFiltered
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} col The index of the column that was automatically filtered.
                 * @param {Array} filterValues The values by which the column was filtered.
                 */
                static RangeFiltered: string;
                /**
                 * Occurs when a column is about to be automatically filtered.
                 * @name GC.Spread.Sheets.Worksheet#RangeFiltering
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} col The index of the column to be automatically filtered.
                 * @param {Array} filterValues The values by which to filter the column.
                 */
                static RangeFiltering: string;
                /**
                 * Occurs when the user has changed the outline state (range group) for rows or columns.
                 * @name GC.Spread.Sheets.Worksheet#RangeGroupStateChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {boolean} isRowGroup Whether the outline (range group) is a group of rows.
                 * @param {number} index The index of the RangeGroupInfo object whose state has changed.
                 * @param {number} level The level of the RangeGroupInfo object whose state has changed.
                 */
                static RangeGroupStateChanged: string;
                /**
                 * Occurs before the user changes the outline state (range group) for rows or columns.
                 * @name GC.Spread.Sheets.Worksheet#RangeGroupStateChanging
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {boolean} isRowGroup Whether the outline (range group) is a group of rows.
                 * @param {number} index The index of the RangeGroupInfo object whose state is changing.
                 * @param {number} level The level of the RangeGroupInfo object whose state is changing.
                 * @param {boolean} cancel A value that indicates whether the operation should be canceled.
                 */
                static RangeGroupStateChanging: string;
                /**
                 * Occurs when a column has just been automatically sorted.
                 * @name GC.Spread.Sheets.Worksheet#RangeSorted
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} col The index of the column that was automatically sorted.
                 * @param {boolean} ascending Whether the automatic sort is ascending.
                 */
                static RangeSorted: string;
                /**
                 * Occurs when a column is about to be automatically sorted.
                 * @name GC.Spread.Sheets.Worksheet#RangeSorting
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} col The index of the column to be automatically sorted.
                 * @param {boolean} ascending Whether the automatic sort is ascending.
                 * @param {boolean} cancel Whether the operation should be canceled.
                 */
                static RangeSorting: string;
                /**
                 * Occurs when a change is made to a row or range of rows in this sheet that may require the row or range of rows to be repainted.
                 * @name GC.Spread.Sheets.Worksheet#RowChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} row The row index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheetArea of the row.
                 * @param {string} propertyName The name of the row's property that has changed.
                 */
                static RowChanged: string;
                /**
                 * Occurs when the row height has changed.
                 * @name GC.Spread.Sheets.Worksheet#RowHeightChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {Array} rowList The list of rows whose heights have changed.
                 * @param {boolean} header Whether the columns are column header columns.
                 */
                static RowHeightChanged: string;
                /**
                 * Occurs when the row height is changing.
                 * @name GC.Spread.Sheets.Worksheet#RowHeightChanging
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {Array} rowList The list of rows whose heights are changing.
                 * @param {boolean} header Whether the columns are column header columns.
                 * @param {boolean} cancel A value that indicates whether the operation should be canceled.
                 */
                static RowHeightChanging: string;
                /**
                 * Occurs when the selection of cells on the sheet has changed.
                 * @name GC.Spread.Sheets.Worksheet#SelectionChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {Array.<GC.Spread.Sheets.Range>} oldSelections The old selection ranges.
                 * @param {Array.<GC.Spread.Sheets.Range>} newSelections The new selection ranges.
                 */
                static SelectionChanged: string;
                /**
                 * Occurs when the selection of cells on the sheet is changing.
                 * @name GC.Spread.Sheets.Worksheet#SelectionChanging
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {Array.<GC.Spread.Sheets.Range>} oldSelections The old selection ranges.
                 * @param {Array.<GC.Spread.Sheets.Range>} newSelections The new selection ranges.
                 */
                static SelectionChanging: string;
                /**
                 * Occurs when the user has changed the sheet name.
                 * @name GC.Spread.Sheets.Worksheet#SheetNameChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} oldValue The sheet's old name.
                 * @param {string} newValue The sheet's new name.
                 */
                static SheetNameChanged: string;
                /**
                 * Occurs when the user is changing the sheet name.
                 * @name GC.Spread.Sheets.Worksheet#SheetNameChanging
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} oldValue The sheet's old name.
                 * @param {string} newValue The sheet's new name.
                 * @param {boolean} cancel A value that indicates whether the operation should be canceled.
                 */
                static SheetNameChanging: string;
                /**
                 * Occurs when the user clicks the sheet tab.
                 * @name GC.Spread.Sheets.Worksheet#SheetTabClick
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} sheetTabIndex The index of the sheet tab that the user clicked.
                 */
                static SheetTabClick: string;
                /**
                 * Occurs when the user double-clicks the sheet tab.
                 * @name GC.Spread.Sheets.Worksheet#SheetTabDoubleClick
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} sheetTabIndex The index of the sheet tab that the user double-clicked.
                 */
                static SheetTabDoubleClick: string;
                /**
                 * Occurs when any slicer has changed.
                 * @name GC.Spread.Sheets.Worksheet#SlicerChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.Slicers.Slicer} slicer The slicer that triggered the event.
                 * @param {string} propertyName The name of the slicer's property that has changed.
                 */
                static SlicerChanged: string;
                /**
                 * Occurs when the sparkline has changed.
                 * @name GC.Spread.Sheets.Worksheet#SparklineChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.Sparklines.Sparkline} sparkline The sparkline whose property has changed.
                 */
                static SparklineChanged: string;
                /**
                 * Occurs when a table column has just been automatically filtered.
                 * @name GC.Spread.Sheets.Worksheet#TableFiltered
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.Tables.Table} table The table column to be automatically filtered.
                 * @param {number} col The index of the table column to be automatically filtered.
                 * @param {Array} filterValues The values by which to filter the column.
                 */
                static TableFiltered: string;
                /**
                 * Occurs when a table column is about to be automatically filtered.
                 * @name GC.Spread.Sheets.Worksheet#TableFiltering
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {GC.Spread.Sheets.Tables.Table} table The table column to be automatically filtered.
                 * @param {number} col The index of the table column to be automatically filtered.
                 * @param {Array} filterValues The values by which to filter the column.
                 */
                static TableFiltering: string;
                /**
                 * Occurs when the top row changes.
                 * @name GC.Spread.Sheets.Worksheet#TopRowChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} oldTopRow The old top row index.
                 * @param {number} newTopRow The new top row index.
                 */
                static TopRowChanged: string;
                /**
                 * Occurs before the touch toolbar pops up.
                 * @name GC.Spread.Sheets.Worksheet#TouchToolStripOpening
                 * @event
                 * @param {number} x The <i>x</i>-coordinate of the horizontal position.
                 * @param {number} y The <i>y</i>-coordinate of the vertical position.
                 * @param {boolean} handled If <c>true</c>, the touch toolbar is prevented from popping up; otherwise, the toolbar is displayed at the default position.
                 */
                static TouchToolStripOpening: string;
                /**
                 * Occurs when the user types a formula.
                 * @name GC.Spread.Sheets.Worksheet#UserFormulaEntered
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} row The row index of the cell in which the user entered a formula.
                 * @param {number} col The column index of the cell in which the user entered a formula.
                 * @param {string} formula The formula that the user entered.
                 */
                static UserFormulaEntered: string;
                /**
                 * Occurs when the user zooms.
                 * @name GC.Spread.Sheets.Worksheet#UserZooming
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} oldZoomFactor The new zoom factor.
                 * @param {number} newZoomFactor The old zoom factor.
                 */
                static UserZooming: string;
                /**
                 * Occurs when the applied cell value is invalid.
                 * @name GC.Spread.Sheets.Worksheet#ValidationError
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} row The cell's row index.
                 * @param {number} col The cell's column index.
                 * @param {GC.Spread.Sheets.DataValidation.DefaultDataValidator} validator The data validator that caused the error.
                 * @param {GC.Spread.Sheets.DataValidation.DataValidationResult} validationResult The policy that the user can set to determine how to process the error.
                 */
                static ValidationError: string;
                /**
                 * Occurs when the value in the subeditor changes.
                 * @name GC.Spread.Sheets.Worksheet#ValueChanged
                 * @event
                 * @param {GC.Spread.Sheets.Worksheet} sheet The sheet that triggered the event.
                 * @param {string} sheetName The sheet's name.
                 * @param {number} row The row index of the cell.
                 * @param {number} col The column index of the cell.
                 * @param {Object} oldValue The old value of the cell.
                 * @param {Object} newValue The new value of the cell.
                 */
                static ValueChanged: string;
            }

            export class LineBorder{
                /**
                 * Represents the line border for a border side.
                 * @class
                 * @param {string} color Indicates the border color and uses a format such as color name (for example, "red") or "#RGB", "#RRGGBB", "rgb(R,B,B)", "rgba(R,G,B,A)".
                 * @param {GC.Spread.Sheets.LineStyle} style Indicates the border line style.
                 */
                constructor(color?:  string,  style?:  GC.Spread.Sheets.LineStyle);
                /**
                 * Indicates the color of the border line. Use a known color name or HEX style color value. The default value is black.
                 */
                color: string;
                /**
                 * Indicates the line style of the border line. The default value is empty.
                 */
                style: GC.Spread.Sheets.LineStyle;
            }

            export class NameInfo{
                /**
                 * Represents a custom named expression that can be used by formulas.
                 * @class
                 * @param {string} name The custom expression name.
                 * @param {Object} expr The custom named expression.
                 * @param {number} row The base row of the expression.
                 * @param {number} column The base column of the expression.
                 */
                constructor(name:  string,  expr:  GC.Spread.CalcEngine.Expression,  row:  number,  column:  number);
                /**
                 * Gets the base column of the custom named expression.
                 * @returns {number} The base column.
                 */
                getColumn(): number;
                /**
                 * Gets the expression.
                 * @returns {GC.Spread.CalcEngine.Expression} The expression.
                 */
                getExpression(): GC.Spread.CalcEngine.Expression;
                /**
                 * Gets the name of the current NameInfo object.
                 * @returns {string} The name of the current NameInfo object.
                 */
                getName(): string;
                /**
                 * Gets the base row of the custom named expression.
                 * @returns {number} The base row.
                 */
                getRow(): number;
            }

            export class Point{
                /**
                 * Represents an <i>x</i>- and <i>y</i>-coordinate pair in two-dimensional space.
                 * @class
                 * @param {number} x The <i>x</i>-coordinate.
                 * @param {number} y The <i>y</i>-coordinate.
                 */
                constructor(x:  number,  y:  number);
                /**
                 * Clones a new point from the current point.
                 * @returns {GC.Spread.Sheets.Point} The cloned object.
                 */
                clone(): GC.Spread.Sheets.Point;
            }

            export class Range{
                /**
                 * Represents a range, which is described by the row index, column index, row count, and column count.
                 * @class
                 * @param {number} r The row index.
                 * @param {number} c The column index.
                 * @param {number} rc The row count.
                 * @param {number} cc The column count.
                 */
                constructor(r:  number,  c:  number,  rc: number,  cc: number);
                /**
                 * The column index.
                 */
                col: number;
                /**
                 * The column count.
                 */
                colCount: number;
                /**
                 * The row index.
                 */
                row: number;
                /**
                 * The row count.
                 */
                rowCount: number;
                /**
                 * Gets whether the current range contains the specified cell.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {number} rowCount The row count.
                 * @param {number} colCount The column count.
                 * @returns {boolean} <c>true</c> if the range contains the cell; otherwise, <c>false</c>.
                 */
                contains(row:  number,  col:  number,  rowCount:  number,  colCount:  number): boolean;
                /**
                 * Gets whether the current range contains the specified range.
                 * @param {GC.Spread.Sheets.Range} range The cell range.
                 * @returns {boolean} <c>true</c> if the current range contains the specified cell range; otherwise, <c>false</c>.
                 */
                containsRange(range:  GC.Spread.Sheets.Range): boolean;
                /**
                 * Gets whether the current range is equal to the specified range.
                 * @param {GC.Spread.Sheets.Range} range The range to compare.
                 * @returns {boolean} <c>true</c> if the current range is equal to the specified range; otherwise, <c>false</c>.
                 */
                equals(range:  GC.Spread.Sheets.Range): boolean;
                /**
                 * Gets the intersection of two cell ranges.
                 * @param {GC.Spread.Sheets.Range} range The cell range.
                 * @param {number} maxRowCount The maximum row count.
                 * @param {number} maxColumnCount The maximum column count.
                 * @returns {GC.Spread.Sheets.Range} Returns null if there is no intersection, or the cell range of the intersection.
                 */
                getIntersect(range:  GC.Spread.Sheets.Range,  maxRowCount:  number,  maxColumnCount:  number): GC.Spread.Sheets.Range;
                /**
                 * Gets whether the current range intersects with the one specified by the row and column index and the row and column count.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {number} rowCount The row count.
                 * @param {number} colCount The column count.
                 * @returns {boolean} <c>true</c> if the specified range intersects with the current range; otherwise <c>false</c>.
                 */
                intersect(row:  number,  col:  number,  rowCount:  number,  colCount:  number): boolean;
                /**
                 * Offsets the location of the range by the specified coordinates.
                 * @param {number} x The offset along the <i>x</i>-axis.
                 * @param {number} y The offset along the <i>y</i>-axis.
                 * @returns {GC.Spread.Sheets.Range} The new location.
                 */
                offset(x:  number,  y:  number): GC.Spread.Sheets.Range;
                /**
                 * Joins this range with the specified range as a union.
                 * @param {GC.Spread.Sheets.Range} range The target range.
                 * @returns {GC.Spread.Sheets.Range} Returns the union of the ranges.
                 */
                union(range:  GC.Spread.Sheets.Range): GC.Spread.Sheets.Range;
            }

            export class Rect{
                /**
                 * Represents a rectangle with a special location, and its width and height in two-dimensional space.
                 * @class
                 * @param {number} x The <i>x</i>-coordinate of the top-left corner of the rectangle.
                 * @param {number} y The <i>y</i>-coordinate of the top-left corner of the rectangle.
                 * @param {number} w The width of the rectangle.
                 * @param {number} h The height of the rectangle.
                 */
                constructor(x:  number,  y:  number,  w:  number,  h:  number);
                /**
                 * The width of the rectangle.
                 */
                height: number;
                /**
                 * The height of the rectangle.
                 */
                width: number;
                /**
                 * The <i>x</i>-coordinate of the top-left corner of the rectangle.
                 */
                x: number;
                /**
                 * The <i>y</i>-coordinate of the top-left corner of the rectangle.
                 */
                y: number;
                /**
                 * Indicates whether the rectangle contains the specified <i>x</i>-coordinate and <i>y</i>-coordinate.
                 * @param {number} x The <i>x</i>-coordinate of the point to check.
                 * @param {number} y The <i>y</i>-coordinate of the point to check.
                 * @returns {boolean} <c>true</c> if (x, y) is contained by the rectangle; otherwise, <c>false</c>.
                 */
                contains(x:  number,  y:  number): boolean;
                /**
                 * Gets the rectangle that intersects with the current rectangle.
                 * @param {GC.Spread.Sheets.Rect} rect The rectangle.
                 * @returns {GC.Spread.Sheets.Rect} The intersecting rectangle. If the two rectangles do not intersect, returns null.
                 */
                getIntersectRect(x:  number,  y:  number,  width:  number,  height:  number): GC.Spread.Sheets.Rect;
                /**
                 * Indicates whether the specified rectangle intersects with the current rectangle.
                 * @param {number} x The <i>x</i>-coordinate of the top-left corner of the rectangle.
                 * @param {number} y The <i>y</i>-coordinate of the top-left corner of the rectangle.
                 * @param {number} w The width of the rectangle.
                 * @param {number} h The height of the rectangle.
                 * @returns {boolean} <c>true</c> if the specified rectangle intersects with the current rectangle; otherwise, <c>false</c>.
                 */
                intersect(x:  number,  y:  number,  width:  number,  height:  number): boolean;
                /**
                 * Indicates whether the specified rectangle intersects with the current rectangle.
                 * @param {GC.Spread.Sheets.Rect} rect The specified rectangle.
                 * @returns {boolean} <c>true</c> if the specified rectangle intersects with the current rectangle; otherwise, <c>false</c>.
                 */
                intersectRect(rect:  GC.Spread.Sheets.Rect): boolean;
            }

            export class Style{
                /**
                 * Represents the style for a cell, row, and column.
                 * @class
                 * @param {string} backColor The background color.
                 * @param {string} foreColor The foreground color.
                 * @param {GC.Spread.Sheets.HorizontalAlign} hAlign The horizontal alignment.
                 * @param {GC.Spread.Sheets.VerticalAlign} vAlign The vertical alignment.
                 * @param {string} font The font.
                 * @param {string} themeFont The font theme.
                 * @param {string|object} formatter The formatting object.
                 * @param {GC.Spread.Sheets.LineBorder} borderLeft The left border.
                 * @param {GC.Spread.Sheets.LineBorder} borderTop The top border.
                 * @param {GC.Spread.Sheets.LineBorder} borderRight The right border.
                 * @param {GC.Spread.Sheets.LineBorder} borderBottom The bottom border.
                 * @param {boolean} locked Whether the cell, row, or column is locked.
                 * @param {number} textIndent The text indent amount.
                 * @param {boolean} wordWrap Whether words wrap within the cell or cells.
                 * @param {boolean} shrinkToFit Whether content shrinks to fit the cell or cells.
                 * @param {string} backgroundImage The background image to display.
                 * @param {GC.Spread.Sheets.CellTypes.Base} cellType The cell type.
                 * @param {GC.Spread.Sheets.ImageLayout} backgroundImageLayout The layout for the background image.
                 * @param {boolean} tabStop Whether the user can set focus to the cell using the Tab key.
                 * @param {GC.Spread.Sheets.TextDecorationType} textDecoration Specifies the decoration added to text.
                 * @param {GC.Spread.Sheets.ImeMode} imeMode Specifies the input method editor mode.
                 * @param {string} name Specifies the name.
                 * @param {string} parentName Specifies the name of the parent style.
                 * @param {string} watermark Specifies the watermark content.
                 * @param {string} cellPadding Specifies the cell padding.
                 * @param {Object} labelOptions Specifies the cell label options.
                 * @param {GC.Spread.Sheets.LabelAlignment} [labelOptions.alignment] The cell label position.
                 * @param {GC.Spread.Sheets.LabelVisibility} [labelOptions.visibility] The cell label visibility.
                 * @param {string} [labelOptions.font] The cell label font.
                 * @param {string} [labelOptions.foreColor] The cell label forecolor.
                 * @param {string} [labelOptions.margin] The cell label margin.
                 */
                constructor(backColor?:  string,  foreColor?:  string,  hAlign?:  GC.Spread.Sheets.HorizontalAlign,  vAlign?:  GC.Spread.Sheets.VerticalAlign,  font?:  any,  themeFont?:  string,  formatter?:  any,  borderLeft?:  GC.Spread.Sheets.LineBorder,  borderTop?:  GC.Spread.Sheets.LineBorder,  borderRight?:  GC.Spread.Sheets.LineBorder,  borderBottom?:  GC.Spread.Sheets.LineBorder,  locked?:  boolean,  textIndent?:  number,  wordWrap?:  boolean,  shrinkToFit?:  boolean,  backgroundImage?:  any,  cellType?:  any,  backgroundImageLayout?:  GC.Spread.Sheets.ImageLayout,  tabStop?:  boolean,  textDecoration?:  GC.Spread.Sheets.TextDecorationType,  imeMode?:  GC.Spread.Sheets.ImeMode,  name?:  string,  parentName?:  string,  watermark?:  string,  cellPadding?:  string,  labelOptions?:  Object);
                /**
                 * Indicates the background color.
                 * @type {string}
                 */
                backColor: string;
                /**
                 * Indicates the background image.
                 * @type {string}
                 */
                backgroundImage: string;
                /**
                 * Indicates the background image layout.
                 * @type {GC.Spread.Sheets.ImageLayout}
                 */
                backgroundImageLayout: GC.Spread.Sheets.ImageLayout;
                /**
                 * Indicates the bottom border line.
                 * @type {GC.Spread.Sheets.LineBorder}
                 */
                borderBottom: GC.Spread.Sheets.LineBorder;
                /**
                 * Indicates the left border line.
                 * @type {GC.Spread.Sheets.LineBorder}
                 */
                borderLeft: GC.Spread.Sheets.LineBorder;
                /**
                 * Indicates the right border line.
                 * @type {GC.Spread.Sheets.LineBorder}
                 */
                borderRight: GC.Spread.Sheets.LineBorder;
                /**
                 * Indicates the top border line.
                 * @type {GC.Spread.Sheets.LineBorder}
                 */
                borderTop: GC.Spread.Sheets.LineBorder;
                /**
                 * Indicates the cell padding.
                 * @type {string}
                 */
                cellPadding: string;
                /**
                 * Indicates the cell type.
                 * @type {GC.Spread.Sheets.CellTypes.Base}
                 */
                cellType: any;
                /**
                 * Indicates the font.
                 * @type {string}
                 */
                font: string;
                /**
                 * Indicates the foreground color.
                 * @type {string}
                 */
                foreColor: string;
                /**
                 * Indicates the formatter.
                 * @type {string|GC.Spread.Formatter.GeneralFormatter}
                 */
                formatter: any;
                /**
                 * Indicates the horizontal alignment.
                 * @type {GC.Spread.Sheets.HorizontalAlign}
                 */
                hAlign: GC.Spread.Sheets.HorizontalAlign;
                /**
                 * Indicates the Input Method Editor (IME) mode.
                 * @type {GC.Spread.Sheets.ImeMode}
                 */
                imeMode: GC.Spread.Sheets.ImeMode;
                /**
                 * Indicates the cell label options.
                 * @property {GC.Spread.Sheets.LabelAlignment} [alignment] - The cell label position.
                 * @property {GC.Spread.Sheets.LabelVisibility} [visibility] - The cell label visibility.
                 * @property {string} [font] - The cell label font.
                 * @property {string} [foreColor] - The cell label forecolor.
                 * @property {string} [margin] - The cell label margin.
                 * @type {Object}
                 */
                labelOptions: Object;
                /**
                 * Indicates whether a cell is marked as locked from editing.
                 * @type {boolean}
                 */
                locked: boolean;
                /**
                 * Indicates the name.
                 * @type {string}
                 */
                name: string;
                /**
                 * Indicates the name of the parent style.
                 * @type {string}
                 */
                parentName: string;
                /**
                 * Indicates whether to shrink to fit.
                 * @type {boolean}
                 */
                shrinkToFit: boolean;
                /**
                 * Indicates whether the user can set focus to the cell using the Tab key.
                 * @type {boolean}
                 */
                tabStop: boolean;
                /**
                 * Indicates the decoration added to text.
                 * @type {GC.Spread.Sheets.TextDecorationType}
                 */
                textDecoration: GC.Spread.Sheets.TextDecorationType;
                /**
                 * Indicates the number of units of indentation for text in a cell, an integer value, where an increment of 1 represents 8 pixels.
                 * @type {number}
                 */
                textIndent: number;
                /**
                 * Indicates the font theme.
                 * @type {string}
                 */
                themeFont: string;
                /**
                 * Indicates the data validator.
                 * @type {GC.Spread.Sheets.DataValidation.DefaultDataValidator}
                 */
                validator: GC.Spread.Sheets.DataValidation.DefaultDataValidator;
                /**
                 * Indicates the vertical alignment.
                 * @type {GC.Spread.Sheets.VerticalAlign}
                 */
                vAlign: GC.Spread.Sheets.VerticalAlign;
                /**
                 * Indicates the watermark content.
                 * @type {string}
                 */
                watermark: string;
                /**
                 * Indicates whether to wrap text.
                 * @type {boolean}
                 */
                wordWrap: boolean;
                /**
                 * Clones the current style.
                 * @returns {GC.Spread.Sheets.Style} The cloned style.
                 */
                clone(): GC.Spread.Sheets.Style;
            }

            export class Theme{
                /**
                 * Represents a color scheme.
                 * @class
                 * @param {string} name The name of the theme.
                 * @param {GC.Spread.Sheets.ColorScheme} colorScheme The base colors of the theme color.
                 * @param {string} headerFont The name of the heading font.
                 * @param {string} bodyFont The name of the body font.
                 */
                constructor(name:  string,  colorScheme:  ColorScheme,  headerFont:  string,  bodyFont:  string);
                /**
                 * Gets or sets the body font of the theme.
                 * @param {string} value The body font.
                 * @returns {string|GC.Spread.Sheets.Theme} If no value is set, returns the body font; otherwise, returns the theme.
                 */
                bodyFont(value?:  string): any;
                /**
                 * Gets or sets the base colors of the theme.
                 * @param {GC.Spread.Sheets.ColorScheme} value The base colors of the theme.
                 * @returns {GC.Spread.Sheets.ColorScheme|GC.Spread.Sheets.Theme} If no value is set, returns the base colors of the theme; otherwise, returns the theme.
                 */
                colors(value?:  GC.Spread.Sheets.ColorScheme): any;
                /**
                 * Gets or sets the heading font of the theme.
                 * @param {string} value The heading font.
                 * @returns {string|GC.Spread.Sheets.Theme} If no value is set, returns the heading font; otherwise, returns the theme.
                 */
                headerFont(value?:  string): any;
                /**
                 * Gets or sets the name of the theme.
                 * @param {string} value The theme name.
                 * @returns {string|GC.Spread.Sheets.Theme} If no value is set, returns the theme name; otherwise, returns the theme.
                 */
                name(value?:  string): any;
            }

            export class ThemeColors{
                /**
                 * Represents the theme color of built-in themes.
                 * @class
                 */
                constructor();
                /**
                 * The theme color of the Apex theme.
                 */
                Apex: ColorScheme;
                /**
                 * The theme color of the Aspect theme.
                 */
                Aspect: ColorScheme;
                /**
                 * The theme color of the Civic theme.
                 */
                Civic: ColorScheme;
                /**
                 * The theme color of the Concourse theme.
                 */
                Concourse: ColorScheme;
                /**
                 * The theme color of the Default theme.
                 */
                Default: ColorScheme;
                /**
                 * The theme color of the Equity theme.
                 */
                Equity: ColorScheme;
                /**
                 * The theme color of the Flow theme.
                 */
                Flow: ColorScheme;
                /**
                 * The theme color of the Foundry theme.
                 */
                Foundry: ColorScheme;
                /**
                 * The theme color of the Median theme.
                 */
                Median: ColorScheme;
                /**
                 * The theme color of the Metro theme.
                 */
                Metro: ColorScheme;
                /**
                 * The theme color of the Module theme.
                 */
                Module: ColorScheme;
                /**
                 * The theme color of the Office theme.
                 */
                Office: ColorScheme;
                /**
                 * The theme color of the Office 2007 theme.
                 */
                Office2007: ColorScheme;
                /**
                 * The theme color of the Opulent theme.
                 */
                Opulent: ColorScheme;
                /**
                 * The theme color of the Oriel theme.
                 */
                Oriel: ColorScheme;
                /**
                 * The theme color of the Origin theme.
                 */
                Origin: ColorScheme;
                /**
                 * The theme color of the Paper theme.
                 */
                Paper: ColorScheme;
                /**
                 * The theme color of the Solstice theme.
                 */
                Solstice: ColorScheme;
                /**
                 * The theme color of the Technic theme.
                 */
                Technic: ColorScheme;
                /**
                 * The theme color of the Trek theme.
                 */
                Trek: ColorScheme;
                /**
                 * The theme color of the Urban theme.
                 */
                Urban: ColorScheme;
                /**
                 * The theme color of the Verve theme.
                 */
                Verve: ColorScheme;
            }

            export class Themes{
                /**
                 * Represents all built-in themes.
                 * @class
                 */
                constructor();
                /**
                 * Indicates the Apex theme.
                 */
                Apex: Theme;
                /**
                 * Indicates the Aspect theme.
                 */
                Aspect: Theme;
                /**
                 * Indicates the Civic theme.
                 */
                Civic: Theme;
                /**
                 * Indicates the Concourse theme.
                 */
                Concourse: Theme;
                /**
                 * Indicates the Default theme.
                 */
                Default: Theme;
                /**
                 * Indicates the Equity theme.
                 */
                Equity: Theme;
                /**
                 * Indicates the Flow theme.
                 */
                Flow: Theme;
                /**
                 * Indicates the Foundry theme.
                 */
                Foundry: Theme;
                /**
                 * Indicates the Median theme.
                 */
                Median: Theme;
                /**
                 * Indicates the Metro theme.
                 */
                Metro: Theme;
                /**
                 * Indicates the Module theme.
                 */
                Module: Theme;
                /**
                 * Indicates the Office theme.
                 */
                Office: Theme;
                /**
                 * Indicates the Office 2007 theme.
                 */
                Office2007: Theme;
                /**
                 * Indicates the Opulent theme.
                 */
                Opulent: Theme;
                /**
                 * Indicates the Oriel theme.
                 */
                Oriel: Theme;
                /**
                 * Indicates the Origin theme.
                 */
                Origin: Theme;
                /**
                 * Indicates the Paper theme.
                 */
                Paper: Theme;
                /**
                 * Indicates the Solstice theme.
                 */
                Solstice: Theme;
                /**
                 * Indicates the Technic theme.
                 */
                Technic: Theme;
                /**
                 * Indicates the Trek theme.
                 */
                Trek: Theme;
                /**
                 * Indicates the Urban theme.
                 */
                Urban: Theme;
                /**
                 * Indicates the Verve theme.
                 */
                Verve: Theme;
            }

            export class Workbook{
                /**
                 * Represents a spreadsheet with the specified hosted DOM element and options setting.
                 * @class
                 * @param {Object} host - The host DOM element.
                 * @param {Object} [options] - The initialization options.<br />
                 * @param {number} [options.sheetCount] - The number of sheets.<br />
                 * @param {string} [options.font] - The tab strip font.<br />
                 * @param {boolean} [options.allowUserDragDrop] - Whether to allow the user to drag and drop range data.<br />
                 * @param {boolean} [options.allowUserDragFill] - Whether to allow the user to drag fill a range.<br />
                 * @param {boolean} [options.allowUserZoom] - Whether to zoom the display by scrolling the mouse wheel while pressing the Ctrl key.<br />
                 * @param {boolean} [options.allowUserResize] - Whether to allow the user to resize columns and rows.<br />
                 * @param {boolean} [options.allowUndo] - Whether to allow the user to undo edits.<br />
                 * @param {boolean} [options.allowSheetReorder] - Whether the user can reorder the sheets in the Spread component.<br />
                 * @param {GC.Spread.Sheets.Fill.AutoFillType} [options.defaultDragFillType] - The default fill type.<br />
                 * @param {boolean} [options.showDragFillSmartTag] - Whether to display the drag fill dialog.<br />
                 * @param {boolean} [options.showHorizontalScrollbar] - Whether to display the horizontal scroll bar.<br />
                 * @param {boolean} [options.showVerticalScrollbar] - Whether to display the vertical scroll bar.<br />
                 * @param {boolean} [options.scrollbarShowMax] - Whether the displayed scroll bars are based on the entire number of columns and rows in the sheet.<br />
                 * @param {boolean} [options.scrollbarMaxAlign] - Whether the scroll bar aligns with the last row and column of the active sheet.<br />
                 * @param {boolean} [options.tabStripVisible] - Whether to display the sheet tab strip.<br />
                 * @param {number} [options.tabStripRatio] - The width of the tab strip expressed as a percentage of the overall horizontal scroll bar width.<br />
                 * @param {boolean} [options.tabEditable] - Whether to allow the user to edit the sheet tab strip.<br />
                 * @param {boolean} [options.newTabVisible] - Whether the spreadsheet displays the special tab to let users insert new sheets.<br />
                 * @param {boolean} [options.tabNavigationVisible] - Whether to display the sheet tab navigation.<br />
                 * @param {boolean} [options.cutCopyIndicatorVisible] - Whether to display an indicator when copying or cutting the selected item.<br />
                 * @param {string} [options.cutCopyIndicatorBorderColor] - The border color for the indicator displayed when the user cuts or copies the selection.<br />
                 * @param {string} [options.backColor] - A color string used to represent the background color of the Spread component, such as "red", "#FFFF00", "rgb(255,0,0)", "Accent 5", and so on.<br />
                 * @param {string} [options.backgroundImage] - The background image of the Spread component.<br />
                 * @param {GC.Spread.Sheets.ImageLayout} [options.backgroundImageLayout] - The background image layout for the Spread component.<br />
                 * @param {string} [options.grayAreaBackColor] - A color string used to represent the background color of the gray area , such as "red", "#FFFF00", "rgb(255,0,0)", "Accent 5", and so on.<br />
                 * @param {GC.Spread.Sheets.ShowResizeTip} [options.showResizeTip] - How to display the resize tip.<br />
                 * @param {boolean} [options.showDragDropTip] -Whether to display the drag-drop tip.<br />
                 * @param {boolean} [options.showDragFillTip] - Whether to display the drag-fill tip.<br />
                 * @param {GC.Spread.Sheets.ShowScrollTip} [options.showScrollTip] - How to display the scroll tip.<br />
                 * @param {boolean} [options.scrollIgnoreHidden] - Whether the scroll bar ignores hidden rows or columns.<br />
                 * @param {boolean} [options.highlightInvalidData] - Whether to highlight invalid data.<br />
                 * @param {boolean} [options.useTouchLayout] - Whether to use touch layout to present the Spread component.<br />
                 * @param {boolean} [options.hideSelection] - Whether to display the selection highlighting when the Spread component does not have focus.<br />
                 * @param {GC.Spread.Sheets.ResizeZeroIndicator} [options.resizeZeroIndicator] - The drawing policy when the row or column is resized to zero.<br />
                 * @param {boolean} [options.allowUserEditFormula] - Whether the user can edit formulas in a cell in the spreadsheet.<br />
                 * @param {boolean} [options.enableFormulaTextbox] - Whether to enable the formula text box in the spreadsheet.<br />
                 * @param {GC.Spread.Sheets.AutoFitType} [options.autoFitType] - Whether content will be formatted to fit in cells or in cells and headers.<br />
                 * @param {GC.Spread.Sheets.ReferenceStyle} [options.referenceStyle] - the style for cell and range references in cell formulas on this sheet.
                 * @param {boolean} [options.allowCopyPasteExcelStyle] - Whether the user can copy style from Spread Sheets then paste to Excel, or copy style from Excel then paste to Spread Sheets.
                 * @param {boolean} [options.allowExtendPasteRange] - Whether extend paste range if the paste range is not enough for pasting.
                 * @param {GC.Spread.Sheets.CopyPasteHeaderOptions} [options.copyPasteHeaderOptions] - Which headers are included when data is copied to or pasted.
                 * @example
                 * var workbook = new GC.Spread.Sheets.Workbook(document.getElementById("ss"), {sheetCount:3, font:"12pt Arial"});
                 * var workbook = new GC.Spread.Sheets.Workbook(document.getElementById("ss"), {sheetcount:3, newTabVisible:false});
                 * var workbook = new GC.Spread.Sheets.Workbook(document.getElementById("ss"), { sheetCount: 3, tabEditable: false });
                 * var workbook = new GC.Spread.Sheets.Workbook(document.getElementById("ss"), {sheetCount:3, tabStripVisible:false});
                 * var workbook = new GC.Spread.Sheets.Workbook(document.getElementById("ss"), {sheetCount:3, allowUserResize:false});
                 * var workbook = new GC.Spread.Sheets.Workbook(document.getElementById("ss"), { sheetCount: 3, allowUserZoom: false});
                 */
                constructor(host?:  any,  options?:  any);
                /**
                 * Represents the name of the Spread control.
                 * @type {string}
                 */
                name: string;
                /**
                 * Represents the options of the Spread control.
                 * @type {Object}
                 * @property {boolean} allowUserDragDrop - Whether to allow the user to drag and drop range data.<br />
                 * @property {boolean} allowUserDragFill - Whether to allow the user to drag fill a range.<br />
                 * @property {boolean} allowUserZoom - Whether to zoom the display by scrolling the mouse wheel while pressing the Ctrl key.<br />
                 * @property {boolean} allowUserResize - Whether to allow the user to resize columns and rows.<br />
                 * @property {boolean} allowUndo - Whether to allow the user to undo edits.<br />
                 * @property {boolean} allowSheetReorder - Whether the user can reorder the sheets in the Spread component.<br />
                 * @property {GC.Spread.Sheets.Fill.AutoFillType} defaultDragFillType - The default fill type.<br />
                 * @property {boolean} showDragFillSmartTag - Whether to display the drag fill dialog.<br />
                 * @property {boolean} showHorizontalScrollbar - Whether to display the horizontal scroll bar.<br />
                 * @property {boolean} showVerticalScrollbar - Whether to display the vertical scroll bar.<br />
                 * @property {boolean} scrollbarShowMax - Whether the displayed scroll bars are based on the entire number of columns and rows in the sheet.<br />
                 * @property {boolean} scrollbarMaxAlign - Whether the scroll bar aligns with the last row and column of the active sheet.<br />
                 * @property {boolean} tabStripVisible - Whether to display the sheet tab strip.<br />
                 * @property {number} tabStripRatio - The width of the tab strip expressed as a percentage of the overall horizontal scroll bar width.<br />
                 * @property {boolean} tabEditable - Whether to allow the user to edit the sheet tab strip.<br />
                 * @property {boolean} newTabVisible - Whether the spreadsheet displays the special tab to let users insert new sheets.<br />
                 * @property {boolean} tabNavigationVisible - Whether to display the sheet tab navigation.<br />
                 * @property {boolean} cutCopyIndicatorVisible - Whether to display an indicator when copying or cutting the selected item.<br />
                 * @property {string} cutCopyIndicatorBorderColor - The border color for the indicator displayed when the user cuts or copies the selection.<br />
                 * @property {string} backColor - A color string used to represent the background color of the Spread component, such as "red", "#FFFF00", "rgb(255,0,0)", "Accent 5", and so on.<br />
                 * @property {string} backgroundImage - The background image of the Spread component.<br />
                 * @property {GC.Spread.Sheets.ImageLayout} backgroundImageLayout - The background image layout for the Spread component.<br />
                 * @property {string} grayAreaBackColor - A color string used to represent the background color of the gray area , such as "red", "#FFFF00", "rgb(255,0,0)", "Accent 5", and so on.<br />
                 * @property {GC.Spread.Sheets.ShowResizeTip} showResizeTip - How to display the resize tip.<br />
                 * @property {boolean} showDragDropTip -Whether to display the drag-drop tip.<br />
                 * @property {boolean} showDragFillTip - Whether to display the drag-fill tip.<br />
                 * @property {GC.Spread.Sheets.ShowScrollTip} showScrollTip - How to display the scroll tip.<br />
                 * @property {boolean} scrollIgnoreHidden - Whether the scroll bar ignores hidden rows or columns.<br />
                 * @property {boolean} highlightInvalidData - Whether to highlight invalid data.<br />
                 * @property {boolean} useTouchLayout - Whether to use touch layout to present the Spread component.<br />
                 * @property {boolean} hideSelection - Whether to display the selection highlighting when the Spread component does not have focus.<br />
                 * @property {GC.Spread.Sheets.ResizeZeroIndicator} resizeZeroIndicator - The drawing policy when the row or column is resized to zero.<br />
                 * @property {boolean} allowUserEditFormula - Whether the user can edit formulas in a cell in the spreadsheet.<br />
                 * @property {boolean} enableFormulaTextbox - Whether to enable the formula text box in the spreadsheet.<br />
                 * @property {GC.Spread.Sheets.AutoFitType} autoFitType - Whether content will be formatted to fit in cells or in cells and headers.<br />
                 * @property {GC.Spread.Sheets.ReferenceStyle} referenceStyle - the style for cell and range references in cell formulas on this sheet.
                 * @property {boolean} allowCopyPasteExcelStyle - Whether the user can copy style from Spread Sheets then paste to Excel, or copy style from Excel then paste to Spread Sheets.
                 * @property {boolean} allowExtendPasteRange - Whether extend paste range if the paste range is not enough for pasting.
                 * @property {GC.Spread.Sheets.CopyPasteHeaderOptions} copyPasteHeaderOptions - Which headers are included when data is copied to or pasted.
                 * @example
                 * var workbook = new GC.Spread.Sheets.Workbook(document.getElementById("ss"),{sheetCount:5,showHorizontalScrollbar:false});
                 * workbook.options.allowUserDragDrop = false;
                 * workbook.options.allowUserZoom = false;
                 */
                options: IWorkbookOptions;
                /**
                 * Represents the sheet collection.
                 * @type {Array.<GC.Spread.Sheets.Worksheet>}
                 */
                sheets: GC.Spread.Sheets.Worksheet[];
                /**
                 * Represents the touch toolstrip.
                 * @type {GC.Spread.Sheets.Touch.TouchToolStrip}
                 */
                touchToolStrip: GC.Spread.Sheets.Touch.TouchToolStrip;
                /**
                 * Adds a custom function.
                 * @param {GC.Spread.CalcEngine.Functions.Function} fn The function to add.
                 */
                addCustomFunction(fn:  GC.Spread.CalcEngine.Functions.Function): void;
                /**
                 * Adds a custom name.
                 * @param {string} name The custom name.
                 * @param {string} formula The formula.
                 * @param {number} baseRow The row index.
                 * @param {number} baseCol The column index.
                 */
                addCustomName(name:  string,  formula:  string,  baseRow:  number,  baseCol:  number): void;
                /**
                 * Adds a style to the Workbook named styles collection.
                 * @param {GC.Spread.Sheets.Style} style The style to be added.
                 */
                addNamedStyle(style:  GC.Spread.Sheets.Style): void;
                /**
                 * Inserts a sheet at the specific index.
                 * @param {number} index The index at which to add a sheet.
                 * @param {number} sheet The sheet to be added.
                 */
                addSheet(index:  number,  sheet?:  GC.Spread.Sheets.Worksheet): void;
                /**
                 * Adds a SparklineEx to the SparklineEx collection.
                 * @param {GC.Spread.Sheets.Sparklines.SparklineEx} sparklineEx The SparklineEx to be added.
                 */
                addSparklineEx(sparklineEx:  GC.Spread.Sheets.Sparklines.SparklineEx): void;
                /**
                 * Binds an event to the Workbook.
                 * @param {string} type The event type.
                 * @param {Object} data Specifies additional data to pass along to the function.
                 * @param {Function} fn Specifies the function to run when the event occurs.
                 */
                bind(type:  any,  data?:  any,  fn?:  any): void;
                /**
                 * Clears all custom functions.
                 */
                clearCustomFunctions(): void;
                /**
                 * Clears custom names.
                 */
                clearCustomNames(): void;
                /**
                 * Clears all sheets in the control.
                 */
                clearSheets(): void;
                /**
                 * Gets the command manager.
                 * @returns {GC.Spread.Commands.CommandManager} The command manager.
                 */
                commandManager(): GC.Spread.Commands.CommandManager;
                /**
                 * Repaints the Workbook control.
                 */
                destroy(): void;
                /**
                 * Makes the Workbook component get focus or lose focus.
                 * @param {boolean} focusIn <c>false</c> makes the Workbook component lose the focus; otherwise, get focus.
                 */
                focus(focusIn?:  boolean): void;
                /**
                 * Loads the object state from the specified JSON string.
                 * @param {Object} workbookData The spreadsheet data from deserialization.
                 */
                fromJSON(workbookData:  Object): void;
                /**
                 * Gets the active sheet.
                 * @returns {GC.Spread.Sheets.Worksheet} The active sheet instance.
                 */
                getActiveSheet(): GC.Spread.Sheets.Worksheet;
                /**
                 * Gets the active sheet index of the control.
                 * @returns {number} The active sheet index.
                 */
                getActiveSheetIndex(): number;
                /**
                 * Gets a custom function.
                 * @param {string} name The custom function name.
                 * @returns {GC.Spread.CalcEngine.Functions.Function} The custom function.
                 */
                getCustomFunction(name:  string): void;
                /**
                 * Gets the specified custom name information.
                 * @param {string} name The custom name.
                 * @returns {GC.Spread.Sheets.NameInfo} The information for the specified custom name.
                 */
                getCustomName(name:  string): GC.Spread.Sheets.NameInfo;
                /**
                 * Gets all custom name information.
                 * @returns {Array} The type GC.Spread.Sheets.NameInfo stored in an array.
                 */
                getCustomNames(): GC.Spread.Sheets.NameInfo[];
                /**
                 * Gets the host element of the current Workbook instance.
                 * @returns {HTMLElement} host The host element of the current Workbook instance.
                 */
                getHost(): HTMLElement;
                /**
                 * Gets a style from the Workbook named styles collection which has the specified name.
                 * @param {string} name The name of the style to return.
                 * @returns {GC.Spread.Sheets.Style} Returns the specified named style.
                 */
                getNamedStyle(name:  string): GC.Spread.Sheets.Style;
                /**
                 * Gets named styles from the Workbook.
                 * @returns {Array} The GC.Spread.Sheets.Style array of named styles.
                 */
                getNamedStyles(): GC.Spread.Sheets.Style[];
                /**
                 * Gets the specified sheet.
                 * @param {number} index The index of the sheet to return.
                 * @returns {GC.Spread.Sheets.Worksheet} The specified sheet.
                 */
                getSheet(index:  number): GC.Spread.Sheets.Worksheet;
                /**
                 * Gets the number of sheets.
                 * @returns {number} The number of sheets.
                 */
                getSheetCount(): number;
                /**
                 * Gets the sheet with the specified name.
                 * @param {string} name The sheet name.
                 * @returns {GC.Spread.Sheets.Worksheet} The sheet with the specified name.
                 */
                getSheetFromName(name:  string): GC.Spread.Sheets.Worksheet;
                /**
                 * Gets the sheet index with the specified name.
                 * @param {string} name The sheet name.
                 * @returns {number} The sheet index.
                 */
                getSheetIndex(name:  string): number;
                /**
                 * Updates the control layout information.
                 */
                invalidateLayout(): void;
                /**
                 * Get if spread paint is suspended.
                 */
                isPaintSuspended(): boolean;
                /**
                 * Gets or sets the next control used by GC.Spread.Sheets.Actions.selectNextControl and GC.Spread.Sheets.Actions.moveToNextCellThenControl.
                 * @param {HTMLElement} value The next control. The control must have a focus method.
                 * @returns {HTMLElement|GC.Spread.Sheets.Workbook} If no value is set, returns the next control; otherwise, returns the spreadsheet.
                 */
                nextControl(value?:  HTMLElement): any;
                /**
                 * Gets or sets the previous control used by GC.Spread.Sheets.Actions.selectPreviousControl and GC.Spread.Sheets.Actions.moveToPreviousCellThenControl.
                 * @param {HTMLElement} value The previous control. The control must have a focus method.
                 * @returns {HTMLElement|GC.Spread.Sheets.Workbook} If no value is set, returns the previous control; otherwise, returns the spreadsheet.
                 */
                previousControl(value?:  HTMLElement): any;
                /**
                 *Prints the specified sheet.
                 *@param {number} sheetIndex The sheet index. If the sheet index is ignored, prints all visible sheets.
                 */
                print(sheetIndex?:  number): void;
                /**
                 * Manually refreshes the layout and rendering of the Workbook object.
                 */
                refresh(): void;
                /**
                 * Removes a custom function.
                 * @param {string} name The custom function name.
                 */
                removeCustomFunction(name:  string): void;
                /**
                 * Removes the specified custom name.
                 * @param {string} name The custom name.
                 */
                removeCustomName(name:  string): void;
                /**
                 * Removes a style from the Workbook named styles collection which has the specified name.
                 * @param {string} name The name of the style to remove.
                 */
                removeNamedStyle(name:  string): void;
                /**
                 * Removes the specified sheet.
                 * @param {number} index The index of the sheet to remove.
                 */
                removeSheet(index:  number): void;
                /**
                 * Removes a SparklineEx from the SparklineEx collection.
                 * @param {string} name The name of the SparklineEx to remove.
                 */
                removeSparklineEx(name:  string): void;
                /**
                 * Repaints the Workbook control.
                 */
                repaint(): void;
                /**
                 * Resumes the calculation service.
                 * @param {boolean} recalcAll Specifies whether to recalculate all formulas.
                 */
                resumeCalcService(ignoreDirty:  boolean): void;
                /**
                 * Resumes the event.
                 */
                resumeEvent(): void;
                /**
                 * Resumes the paint of active sheet and tab strip.
                 */
                resumePaint(): void;
                /**
                 * Searches the text in the cells in the specified sheet for the specified string with the specified criteria.
                 * @param {GC.Spread.Sheets.Search.SearchCondition} searchCondition The search conditions.
                 * @returns {GC.Spread.Sheets.Search.SearchResult} The search result.
                 */
                search(searchCondition:  GC.Spread.Sheets.Search.SearchCondition): GC.Spread.Sheets.Search.SearchResult;
                /**
                 * Sets the active sheet by name.
                 * @param {string} name The name of the sheet to make the active sheet.
                 */
                setActiveSheet(name:  string): void;
                /**
                 * Sets the active sheet index for the control.
                 * @param {number} value The active sheet index.
                 */
                setActiveSheetIndex(value:  number): void;
                /**
                 * Sets the number of sheets.
                 * @param {number} count The number of sheets.
                 */
                setSheetCount(count:  number): void;
                /**
                 * Gets or sets the index of the first sheet to display in the spreadsheet.
                 * @param {number} value The index of the first sheet to display in the spreadsheet.
                 * @returns {number|GC.Spread.Sheets.Workbook} If no value is set, returns the index of the first sheet displayed in the spreadsheet; otherwise, returns the spreadsheet.
                 */
                startSheetIndex(value?:  number): any;
                /**
                 * Suspends the calculation service.
                 * @param {boolean} ignoreDirty Specifies whether to invalidate the dependency cells.
                 */
                suspendCalcService(ignoreDirty:  boolean): void;
                /**
                 * Suspends the event.
                 */
                suspendEvent(): void;
                /**
                 * Suspends the paint of active sheet and tab strip.
                 */
                suspendPaint(): void;
                /**
                 * Saves the object state to a JSON string.
                 * @param {Object} serializationOption Serialization option that contains the <i>includeBindingSource</i> argument. See the Remarks for more information.
                 * @returns {Object} The spreadsheet data.
                 */
                toJSON(serializationOption:  Object): Object;
                /**
                 * Removes the binding of an event to Workbook.
                 * @param {string} type The event type.
                 * @param {Function} fn Specifies the function to run when the event occurs.
                 */
                unbind(type:  any,  fn?:  any): void;
                /**
                 * Removes the binding of all events to Workbook.
                 */
                unbindAll(): void;
                /**
                 * Gets the undo manager.
                 * @returns {GC.Spread.Commands.UndoManager} The undo manager.
                 */
                undoManager(): GC.Spread.Commands.UndoManager;
            }

            export class Worksheet{
                /**
                 * Represents a worksheet.
                 * @class
                 * @param {string} name The name of the Worksheet.
                 */
                constructor(name:  string);
                /**
                 * Indicates whether to generate columns automatically while binding data context.
                 * @type {boolean}
                 */
                autoGenerateColumns: boolean;
                /**
                 * Indicates the column range group.
                 * @type {GC.Spread.Sheets.Outlines.Outline}
                 */
                columnOutlines: GC.Spread.Sheets.Outlines.Outline;
                /**
                 * Comment manager for the sheet.
                 * @type {GC.Spread.Sheets.Comments.CommentManager}
                 */
                comments: GC.Spread.Sheets.Comments.CommentManager;
                /**
                 * Conditional format manager for the sheet.
                 * @type {GC.Spread.Sheets.ConditionalFormatting.ConditionalFormats}
                 */
                conditionalFormats: ConditionalFormatting.ConditionalFormats;
                /**
                 * Indicates the default row height and column width of the sheet.
                 * @type {Object}
                 */
                defaults: GC.Spread.Sheets.ISheetDefaultOption;
                /**
                 * FloatingObject manager for the sheet.
                 * @type {GC.Spread.Sheets.FloatingObjects.FloatingObjectCollection}
                 */
                floatingObjects: GC.Spread.Sheets.FloatingObjects.FloatingObjectCollection;
                /**
                 * Indicates the options of the sheet.
                 * @type {Object}
                 * @property {boolean} allowCellOverflowindicates - whether data can overflow into adjacent empty cells.
                 * @property {string} sheetTabColor - A color string used to represent the sheet tab color, such as "red", "#FFFF00", "rgb(255,0,0)", "Accent 5", and so on.
                 * @property {string} frozenlineColor - A color string used to represent the frozen line color, such as "red", "#FFFF00", "rgb(255,0,0)", "Accent 5", and so on.
                 * @property {GC.Spread.Sheets.ClipboardPasteOptions} clipBoardOptions - The clipboard option.
                 * @property {Object} gridline - The grid line's options.
                 * @property {string} gridline.color - The grid line color
                 * @property {boolean} gridline.showVerticalGridline - Whether to show the vertical grid line.
                 * @property {boolean} gridline.showHorizontalGridline - Whether to show the horizontal grid line.
                 * @property {boolean} rowHeaderVisible - Indicates whether the row header is visible.
                 * @property {boolean} colHeaderVisible - Indicates whether the column header is visible.
                 * @property {GC.Spread.Sheets.HeaderAutoText} rowHeaderAutoText - Indicates whether the row header displays letters or numbers or is blank.
                 * @property {GC.Spread.Sheets.HeaderAutoText} colHeaderAutoText - Indicates whether the column header displays letters or numbers or is blank.
                 * @property {number} rowHeaderAutoTextIndex - Specifies which row header column displays the automatic text when there are multiple row header columns.
                 * @property {number} colHeaderAutoTextIndex - Specifies which column header row displays the automatic text when there are multiple column header rows.
                 * @property {boolean} isProtected - Indicates whether cells on this sheet that are marked as protected cannot be edited.
                 * @property {Object} protectionOptions - A value that indicates the elements that you want users to be able to change.
                 * @property {boolean} [protectionOptions.allowSelectLockedCells] - True or undefined if the user can select locked cells.
                 * @property {boolean} [protectionOptions.allowSelectUnlockedCells] - True or undefined if the user can select unlocked cells.
                 * @property {boolean} [protectionOptions.allowSort] - True if the user can sort ranges.
                 * @property {boolean} [protectionOptions.allowFilter] - True if the user can filter ranges.
                 * @property {boolean} [protectionOptions.allowEditObjects] - True if the user can edit floating objects.
                 * @property {boolean} [protectionOptions.allowResizeRows] - True if the user can resize rows.
                 * @property {boolean} [protectionOptions.allowResizeColumns] - True if the user can resize columns.
                 * @property {string} selectionBackColor - The selection's background color for the sheet.
                 * @property {string} selectionBorderColor -  The selection's border color for the sheet.
                 */
                options: GC.Spread.Sheets.IWorksheetOptions;
                /**
                 * Picture manager for the sheet.
                 * @type {GC.Spread.Sheets.FloatingObjects.FloatingObjectCollection}
                 */
                pictures: GC.Spread.Sheets.FloatingObjects.FloatingObjectCollection;
                /**
                 * Indicates the row range group.
                 * @type {GC.Spread.Sheets.Outlines.Outline}
                 */
                rowOutlines: GC.Spread.Sheets.Outlines.Outline;
                /** The slicer manager.
                 * @type {GC.Spread.Sheets.Slicers.SlicerCollection}
                 */
                slicers: GC.Spread.Sheets.Slicers.SlicerCollection;
                /** The table manager.
                 * @type {GC.Spread.Sheets.Tables.TableManager}
                 */
                tables: GC.Spread.Sheets.Tables.TableManager;
                /**
                 * Adds the column or columns to the data model at the specified index.
                 * @param {number} col Column index at which to add the new columns.
                 * @param {number} count The number of columns to add.
                 */
                addColumns(col:  number,  count:  number): void;
                /**
                 * Adds a custom function.
                 * @param {GC.Spread.CalcEngine.Functions.Function} fn The function to add.
                 */
                addCustomFunction(fn:  GC.Spread.CalcEngine.Functions.Function): void;
                /**
                 * Adds a custom name.
                 * @param {string} name The custom name.
                 * @param {string} formula The formula.
                 * @param {number} baseRow The row index.
                 * @param {number} baseCol The column index.
                 */
                addCustomName(name:  string,  formula:  string,  baseRow:  number,  baseCol:  number): void;
                /**
                 * Adds a style to the Worksheet named styles collection.
                 * @param {GC.Spread.Sheets.Style} style The style to be added.
                 */
                addNamedStyle(style:  GC.Spread.Sheets.Style): void;
                /**
                 * Adds rows in this worksheet.
                 * @param {number} row The index of the starting row.
                 * @param {number} count The number of rows to add.
                 */
                addRows(row:  number,  count:  number): void;
                /**
                 * Adds a cell or cells to the selection.
                 * @param {number} row The row index of the first cell to add.
                 * @param {number} column The column index of the first cell to add.
                 * @param {number} rowCount The number of rows to add.
                 * @param {number} columnCount The number of columns to add.
                 */
                addSelection(row:  number,  column:  number,  rowCount:  number,  columnCount:  number): void;
                /**
                 * Adds a span of cells to this sheet in the specified sheet area.
                 * @param {number} row The row index of the cell at which to start the span.
                 * @param {number} column The column index of the cell at which to start the span.
                 * @param {number} rowCount The number of rows to span.
                 * @param {number} colCount The number of columns to span.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 */
                addSpan(row:  number,  col:  number,  rowCount:  number,  colCount:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Automatically fits the viewport column.
                 * @param {number} column The column index.
                 */
                autoFitColumn(column:  number): void;
                /**
                 * Automatically fits the viewport row.
                 * @param {number} row The row index.
                 */
                autoFitRow(row:  number): void;
                /**
                 * Binds an event to the sheet.
                 * @param {string} type The event type.
                 * @param {Object} data Optional. Specifies additional data to pass along to the function.
                 * @param {Function} fn Specifies the function to run when the event occurs.
                 */
                bind(type:  any,  data?:  any,  fn?:  any): void;
                /**
                 * Binds the column using the specified data field.
                 * @param {number} index The column index.
                 * @param {string|Object} column Column information with data field. If its type is string, it is regarded as name.
                 */
                bindColumn(index:  number,  column:  any): void;
                /**
                 * Binds the columns using the specified data fields.
                 * @param {Array} columns The array of column information with data fields. If an item's type is string, the item is regarded as name.
                 */
                bindColumns(columns:  any[]): void;
                /**
                 * Clears the specified area.
                 * @param {number} row The start row index.
                 * @param {number} column The start column index.
                 * @param {number} rowCount The number of rows to clear.
                 * @param {number} columnCount The number of columns to clear.
                 * @param {GC.Spread.Sheets.SheetArea} area The area to clear.
                 * @param {GC.Spread.Sheets.StorageType} storageType The clear type.
                 */
                clear(row:  number,  column:  number,  rowCount:  number,  colCount:  number,  area:  GC.Spread.Sheets.SheetArea,  storageType:  GC.Spread.Sheets.StorageType): void;
                /**
                 * Clears all custom functions.
                 */
                clearCustomFunctions(): void;
                /**
                 * Clears custom names.
                 */
                clearCustomNames(): void;
                /**
                 * Clears the dirty, insert, and delete status from the current worksheet.
                 */
                clearPendingChanges(): void;
                /**
                 * Clears the selection.
                 */
                clearSelection(): any;
                /**
                 * Copies data from one range to another.
                 * @param {number} fromRow The source row.
                 * @param {number} fromColumn The source column.
                 * @param {number} toRow The target row.
                 * @param {number} toColumn The target column.
                 * @param {number} rowCount The row count.
                 * @param {number} columnCount The column count.
                 * @param {GC.Spread.Sheets.CopyToOptions} option The copy option.
                 */
                copyTo(fromRow:  number,  fromColumn:  number,  toRow:  number,  toColumn:  number,  rowCount:  number,  columnCount:  number,  option:  GC.Spread.Sheets.CopyToOptions): void;
                /**
                 * Gets or sets the current theme for the sheet.
                 * @param {string|GC.Spread.Common.Theme} value The theme name or the theme.
                 * @returns {GC.Spread.Common.Theme|GC.Spread.Sheets.Worksheet} If no value is set, returns the current theme; otherwise, returns the worksheet.
                 */
                currentTheme(value?:  any): any;
                /**
                 * Deletes the columns in this sheet at the specified index.
                 * @param {number} col The index of the first column to delete.
                 * @param {number} count The number of columns to delete.
                 */
                deleteColumns(col:  number,  count:  number): void;
                /**
                 * Deletes the rows in this worksheet at the specified index.
                 * @param {number} row The index of the first row to delete.
                 * @param {number} count The number of rows to delete.
                 */
                deleteRows(row:  number,  count:  number): void;
                /**
                 * Returns the editor's status.
                 * @returns {GC.Spread.Sheets.EditorStatus} The editor status.
                 */
                editorStatus(): GC.Spread.Sheets.EditorStatus;
                /**
                 * Stops editing the active cell.
                 * @param {boolean} ignoreValueChange If set to <c>true</c>, does not apply the edited text to the cell.
                 * @returns {boolean} <c>true</c> when able to stop cell editing successfully; otherwise, <c>false</c>.
                 */
                endEdit(ignoreValueChange?:  boolean): boolean;
                /**
                 * Fills the specified range automatically.
                 * @param {GC.Spread.Sheets.Range} startRange The fill start range.
                 * @param {GC.Spread.Sheets.Range} wholeRange The entire range to fill.
                 * @param {object} options The range fill information.<br />
                 * options.fillType {GC.Spread.Sheets.Fill.FillType} fillType Specifies how to fill the specified range.<br />
                 *      GC.Spread.Sheets.Fill.FillType.direction:<br />
                 *              Fills the specified range in the specified direction.<br />
                 *      GC.Spread.Sheets.Fill.FillType.linear:<br />
                 *              Fills the specified range using a linear trend when the source value type is number.<br />
                 *              The next value is generated by the step and stop values.<br />
                 *              The next value is computed by adding the step value to the current cell value.<br />
                 *      GC.Spread.Sheets.Fill.FillType.growth:<br />
                 *              Fills the specified range using a growth trend when the source value type is number.<br />
                 *              The next value is generated by the step and stop values.<br />
                 *              The next value is computed by multiplying the step value with the current cell.<br />
                 *      GC.Spread.Sheets.Fill.FillType.date:<br />
                 *              Fills the specified range when the source value type is date.<br />
                 *              The next value is generated by adding the step value to the current value.<br />
                 *              The step value is affected by the fill date unit.<br />
                 *      GC.Spread.Sheets.Fill.FillType.auto:<br />
                 *              Fills the specified range automatically.<br />
                 *              When the value is a string, the value is copied to other cells.<br />
                 *              When the value is a number, the new value is generated by the TREND formula.<br />
                 * options.series {GC.Spread.Sheets.Fill.FillSeries} series The fill series.<br />
                 * options.direction {GC.Spread.Sheets.Fill.FillDirection} direction The fill direction.<br />
                 * options.step {number} step The fill step value.<br />
                 * options.stop {number|Date} stop The fill stop value.<br />
                 * options.unit {GC.Spread.Sheets.Fill.FillDateUnit} unit The fill date unit.
                 */
                fillAuto(startRange:  GC.Spread.Sheets.Range,  wholeRange:  GC.Spread.Sheets.Range,  options:  Object): void;
                /**
                 * Loads the object state from the specified JSON string.
                 * @param {Object} sheetSettings The sheet data from deserialization.
                 */
                fromJSON(sheetSettings:  Object,  setFormulaDirectly?:  boolean,  noSchema?:  boolean): void;
                /**
                 * Gets or sets the number of frozen columns of the sheet.
                 * @param {number} colCount The number of columns to freeze.
                 * @returns {number|GC.Spread.Sheets.Worksheet} If no value is set, returns the number of frozen columns; otherwise, returns the worksheet.
                 */
                frozenColumnCount(colCount?:  number): any;
                /**
                 * Gets or sets the number of frozen rows of the sheet.
                 * @param {number} rowCount The number of rows to freeze.
                 * @returns {number|GC.Spread.Sheets.Worksheet} If no value is set, returns the number of frozen rows; otherwise, returns the worksheet.
                 */
                frozenRowCount(rowCount?:  number): any;
                /**
                 * Gets or sets the number of trailing frozen columns of the sheet.
                 * @param {number} colCount The number of columns to freeze at the right side of the sheet.
                 * @returns {number|GC.Spread.Sheets.Worksheet} If no value is set, returns the number of trailing frozen columns; otherwise, returns the worksheet.
                 */
                frozenTrailingColumnCount(colCount?:  number): any;
                /**
                 * Gets or sets the number of trailing frozen rows of the sheet.
                 * @param {number} rowCount The number of rows to freeze at the bottom of the sheet.
                 * @returns {number|GC.Spread.Sheets.Worksheet} If no value is set, returns the number of trailing frozen rows; otherwise, returns the worksheet.
                 */
                frozenTrailingRowCount(rowCount?:  number): any;
                /**
                 * Gets the active column index for this sheet.
                 * @returns {number} The column index of the active cell.
                 */
                getActiveColumnIndex(): number;
                /**
                 * Gets the active row index for this sheet.
                 * @returns {number} The row index of the active cell.
                 */
                getActiveRowIndex(): number;
                /**
                 * Gets the actual style information for a specified cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} column The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 * @param {boolean} sheetStyleOnly If <c>true</c>, the row filter and the conditional format style are not applied to the return style;
                 * otherwise, the return style only contains the cell's inherited style.
                 * @returns {GC.Spread.Sheets.Style} Returns the cell style of the specified cell.
                 */
                getActualStyle(row:  number,  column:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea,  sheetStyleOnly?:  boolean): GC.Spread.Sheets.Style;
                /**
                 * Gets an object array from a specified range of cells.
                 * @param {number} row The row index.
                 * @param {number} column The column index.
                 * @param {number} rowCount The row count.
                 * @param {number} colCount The column count.
                 * @param {boolean} getFormula If <c>true</c>, return formulas; otherwise, return values.
                 * @returns {Array} The object array from the specified range of cells.
                 */
                getArray(row:  number,  column:  number,  rowCount:  number,  columnCount:  number,  getFormula?:  boolean): any[];
                /**
                 * Gets the binding path of cell-level binding from the specified cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @returns {string} Returns the binding path of the cell for cell-level binding.
                 */
                getBindingPath(row:  number,  col:  number): string;
                /**
                 * Gets the specified cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 * @returns {GC.Spread.Sheets.CellRange} The cell.
                 */
                getCell(row:  number,  col:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): GC.Spread.Sheets.CellRange;
                /**
                 * Gets the rectangle of the cell.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {number} rowViewportIndex Index of the row of the viewport: -1 represents column header area, 0 represents frozen row area, 1 represents viewport area, 2 represents trailing frozen row area.
                 * @param {number} colViewportIndex Index of the column of the viewport: -1 represents row header area, 0 represents frozen column area, 1 represents viewport area, 2 represents trailing frozen column area.
                 * @returns {GC.Spread.Sheets.Rect} Object that contains the size and location of the cell rectangle.
                 */
                getCellRect(row:  number,  col:  number,  rowViewportIndex?:  number,  colViewportIndex?:  number): GC.Spread.Sheets.Rect;
                /**
                 * Gets the cell type.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 * @returns {GC.Spread.Sheets.CellTypes.Base} Returns the cell type for the specified cell.
                 */
                getCellType(row:  number,  col:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): GC.Spread.Sheets.CellTypes.Base;
                /**
                 * Gets the column count in the specified sheet area.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 * @returns {number} The number of columns.
                 */
                getColumnCount(sheetArea?:  GC.Spread.Sheets.SheetArea): number;
                /**
                 * Gets whether a forced page break is inserted before the specified column on this sheet when printing.
                 * @param {number} column The column index.
                 * @returns {boolean} <c>true</c> if a forced page break is inserted before the specified column; otherwise, <c>false</c>.
                 */
                getColumnPageBreak(column:  number): boolean;
                /**
                 * Gets a value that indicates whether the user can resize a specified column in the specified sheet area.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 * @returns {boolean} <c>true</c> if the user can resize the specified column; otherwise, <c>false</c>.
                 */
                getColumnResizable(col:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): boolean;
                /**
                 * Gets whether a column in the specified sheet area is displayed.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 * @returns {boolean} <c>true</c> if the column is visible in the sheet area; otherwise, <c>false</c>.
                 */
                getColumnVisible(col:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): boolean;
                /**
                 * Gets the width in pixels for the specified column in the specified sheet area.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to viewport.
                 * @returns {number} The column width in pixels.
                 */
                getColumnWidth(col:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): number;
                /**
                 * Gets delimited text from a range.
                 * @param {number} row The start row.
                 * @param {number} column The start column.
                 * @param {number} rowCount The row count.
                 * @param {number} columnCount The column count.
                 * @param {string} rowDelimiter The row delimiter that is appended to the end of the row.
                 * @param {string} columnDelimiter The column delimiter that is appended to the end of the column.
                 * @returns {string} The text from the range with the specified delimiters.
                 */
                getCsv(row:  number,  column:  number,  rowCount:  number,  columnCount:  number,  rowDelimiter:  string,  columnDelimiter:  string): string;
                /**
                 * Gets a custom function.
                 * @param {string} fnName The custom function name.
                 * @returns {GC.Spread.CalcEngine.Functions.Function} The custom function.
                 */
                getCustomFunction(name:  string): void;
                /**
                 * Gets the specified custom name information.
                 * @param {string} fnName The custom name.
                 * @returns {GC.Spread.Sheets.NameInfo} The information for the specified custom name.
                 */
                getCustomName(name:  string): GC.Spread.Sheets.NameInfo;
                /**
                 * Gets all custom name information.
                 * @returns {Array} The type GC.Spread.Sheets.NameInfo stored in an array.
                 */
                getCustomNames(): GC.Spread.Sheets.NameInfo[];
                /**
                 * Gets the column name at the specified position.
                 * @param {number} column The column index for which the name is requested.
                 * @returns {string} The column name for data binding.
                 */
                getDataColumnName(column:  number): string;
                /**
                 * Gets the data item.
                 * @param {number} row The row index.
                 * @returns {Object} The row data.
                 */
                getDataItem(row:  number): any;
                /**
                 * Gets the data source that populates the sheet.
                 * @function
                 * @returns {Object} Returns the data source.
                 */
                getDataSource(): any;
                /**
                 * Gets the cell data validator.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 * @returns {GC.Spread.Sheets.DataValidation.DefaultDataValidator} Returns the cell data validator for the specified cell.
                 */
                getDataValidator(row:  number,  col:  number,  sheetArea:  GC.Spread.Sheets.SheetArea): GC.Spread.Sheets.DataValidation.DefaultDataValidator;
                /**
                 * Gets the default style information for the sheet.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 * @returns {GC.Spread.Sheets.Style} Returns the sheet's default style.
                 */
                getDefaultStyle(sheetArea?:  GC.Spread.Sheets.SheetArea): GC.Spread.Sheets.Style;
                /**
                 * Gets the deleted row collection.
                 * @return {Array} The deleted rows.
                 */
                getDeletedRows(): any[];
                /**
                 * Gets the dirty cell collection.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {number} rowCount The number of rows in the range of dirty cells.
                 * @param {number} colCount The number of columns in the range of dirty cells.
                 * @return {Array} The dirty cells.
                 */
                getDirtyCells(row:  number,  col:  number,  rowCount:  number,  colCount:  number): GC.Spread.Sheets.IDirtyCellInfo[];
                /**
                 * Gets the dirty row collection.
                 * @returns {Array} The dirty rows.
                 */
                getDirtyRows(): any[];
                /**
                 * Gets the cell formatter.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 * @returns {Object} Returns the cell formatter string or object for the specified cell.
                 */
                getFormatter(row:  number,  col:  number,  sheetArea:  GC.Spread.Sheets.SheetArea): Object;
                /**
                 * Gets the formula in the specified cell in this sheet.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @returns {string} Returns the formula string.
                 */
                getFormula(row:  number,  col:  number): string;
                /**
                 * Gets the formula in the specified cell in this sheet.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @returns {object} Returns the formula information about the cell.<br />
                 * object.hasFormula {boolean} Indicates whether there is a formula in the cell.<br />
                 * object.formula {string} The formula string.<br />
                 * object.isArrayFormula {boolean} Indicates whether the formula is an array formula.<br />
                 * object.baseRange {GC.Spread.Sheets.Range} The range applied to the array formula.
                 */
                getFormulaInformation(row:  number,  col:  number): IFormulaInformation;
                /**
                 * Gets the inserted row collection.
                 * @returns {Array} The inserted rows.
                 */
                getInsertRows(): any[];
                /**
                 * Gets a style from the Worksheet named styles collection which has the specified name.
                 * @param {string} name The name of the style to return.
                 * @returns {GC.Spread.Sheets.Style} Returns the specified named style.
                 */
                getNamedStyle(name:  string): GC.Spread.Sheets.Style;
                /**
                 * Gets named styles from the Worksheet.
                 * @returns {Array} The GC.Spread.Sheets.Style array of named styles.
                 */
                getNamedStyles(): GC.Spread.Sheets.Style[];
                /**
                 * Gets the parent Spread object of the current sheet.
                 * @returns {GC.Spread.Sheets.Workbook} Returns the parent Spread object of the current sheet.
                 */
                getParent(): GC.Spread.Sheets.Workbook;
                /**
                 * Gets a range of cells in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {number} rowCount The row count of the range. If you do not provide this parameter, it defaults to <b>1</b>.
                 * @param {number} colCount The column count of the range. If you do not provide this parameter, it defaults to <b>1</b>.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 * @returns {GC.Spread.Sheets.CellRange} The cellRange.
                 * If row is -1 and rowCount is -1, the range represents columns. For example, sheet.getRange(-1,4,-1,6) returns the columns "E:J".
                 * If col is -1 and colCount is -1, the range represents rows. For example, sheet.getRange(4,-1,6,-1) returns the rows "5:10".
                 */
                getRange(row:  number,  col:  number,  rowCount?:  number,  colCount?:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): GC.Spread.Sheets.CellRange;
                /**
                 * Gets the row count in the specified sheet area.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 * @returns {number} The number of rows.
                 */
                getRowCount(sheetArea?:  GC.Spread.Sheets.SheetArea): number;
                /**
                 * Gets the height in pixels for the specified row in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 * @returns {number} The row height in pixels.
                 */
                getRowHeight(row:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): number;
                /**
                 * Gets whether a forced page break is inserted before the specified row on this sheet when printing.
                 * @param {number} row The row index.
                 * @returns {boolean} <c>true</c> if a forced page break is inserted before the specified row; otherwise, <c>false</c>.
                 */
                getRowPageBreak(row:  number): boolean;
                /**
                 * Gets a value that indicates whether users can resize the specified row in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 * @returns {boolean} <c>true</c> if the users can resize the specified row; otherwise, <c>false</c>.
                 */
                getRowResizable(row:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): boolean;
                /**
                 * Gets whether the control displays the specified row.
                 * @param {number} row The row index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 * @returns {boolean} <c>true</c> if the row is visible in the sheet area; otherwise, <c>false</c>.
                 */
                getRowVisible(row:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): boolean;
                /**
                 * Gets the selections in the current sheet.
                 * @returns {Array.<GC.Spread.Sheets.Range>} The type GC.Spread.Sheets.Range is stored in an Array.
                 */
                getSelections(): GC.Spread.Sheets.Range[];
                /**
                 * Gets the spans in the specified range in the specified sheet area.
                 * @param {GC.Spread.Sheets.Range} range The cell range.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 * @returns {Array} An array that contains span information whose item type is GC.Spread.Sheets.Range.
                 */
                getSpans(range:  GC.Spread.Sheets.Range): GC.Spread.Sheets.Range[];
                /**
                 *  Gets the sparkline for the specified cell.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @returns {GC.Spread.Sheets.Sparkline} The sparkline for the cell.
                 */
                getSparkline(row:  number,  column:  number): GC.Spread.Sheets.Sparklines.Sparkline;
                /**
                 * Gets the style information for a specified cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} column The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 * @returns {GC.Spread.Sheets.Style} Returns the cell style of the specified cell.
                 */
                getStyle(row:  number,  column:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): GC.Spread.Sheets.Style;
                /**
                 * Gets the name of the style for a specified cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} column The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 * @returns {string} Returns the name string for the style.
                 */
                getStyleName(row:  number,  column:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): string;
                /**
                 * Gets the tag value from the specified cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 * @returns {Object} Returns the tag value of the cell.
                 */
                getTag(row:  number,  col:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): any;
                /**
                 * Gets the formatted text in the cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 * @returns {string} Returns the formatted text of the cell.
                 */
                getText(row:  number,  col:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): string;
                /**
                 * Gets the unformatted data from the specified cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 * @returns {Object} Returns the value of the cell.
                 */
                getValue(row:  number,  col:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): any;
                /**
                 * Gets the index of the bottom row in the viewport.
                 * @param {number} rowViewportIndex The index of the viewport.
                 * @returns {number} The index of the bottom row in the viewport.
                 */
                getViewportBottomRow(rowViewportIndex:  number): number;
                /**
                 * Gets the height of the specified viewport row for the active sheet.
                 * @param {number} rowViewportIndex The index of the row viewport.
                 * @returns {number} The height of the viewport.
                 */
                getViewportHeight(rowViewportIndex:  number): number;
                /**
                 * Gets the index of the left column in the viewport.
                 * @param {number} columnViewportIndex The index of the viewport.
                 * @returns {number} The index of the left column in the viewport.
                 */
                getViewportLeftColumn(columnViewportIndex:  number): number;
                /**
                 * Gets the index of the right column in the viewport.
                 * @param {number} columnViewportIndex The index of the viewport.
                 * @returns {number} The index of the right column in the viewport.
                 */
                getViewportRightColumn(columnViewportIndex:  number): number;
                /**
                 * Gets the index of the top row in the viewport.
                 * @param {number} rowViewportIndex The index of the viewport.
                 * @returns {number} The index of the top row in the viewport.
                 */
                getViewportTopRow(rowViewportIndex:  number): number;
                /**
                 * Gets the width of the specified viewport column for the active sheet.
                 * @param {number} columnViewportIndex The index of the column viewport.
                 * @returns {number} The width of the viewport
                 */
                getViewportWidth(columnViewportIndex:  number): number;
                /**
                 * Groups the sparklines.
                 * @param {Array.<GC.Spread.Sheets.Sparklines.Sparkline>} sparklines The sparklines to group.
                 * @returns {GC.Spread.Sheets.Sparklines.SparklineGroup} The sparkline group.
                 */
                groupSparkline(sparklines:  Sparklines.Sparkline[]): GC.Spread.Sheets.Sparklines.SparklineGroup;
                /**
                 * Gets whether there is a dirty, insert, or delete status for the specified range.
                 * @returns {boolean} <c>true</c> if any of the rows or cells in the range are dirty, or have been inserted or deleted; otherwise, <c>false</c>.
                 */
                hasPendingChanges(): boolean;
                /**
                 * Performs a hit test.
                 * @param {number} x The <i>x</i>-coordinate.
                 * @param {number} y The <i>y</i>-coordinate.
                 * @returns {Object} The hit test information.
                 */
                hitTest(x:  number,  y:  number): IHitTestInformation;
                /**
                 * Invalidates the sheet layout.
                 */
                invalidateLayout(): void;
                /**
                 * Gets whether the sheet is in edit mode.
                 * @returns {boolean} <c>true</c> if the sheet is in edit mode; otherwise, <c>false</c>.
                 */
                isEditing(): boolean;
                /**
                 * Get if sheet paint is suspended.
                 */
                isPaintSuspended(): boolean;
                /**
                 * Determines whether the cell value is valid.
                 * @param {number} row The row index.
                 * @param {number} column The column index.
                 * @param {Object} value The cell value.
                 * @returns {boolean} <c>true</c> if the value is valid; otherwise, <c>false</c>.
                 */
                isValid(row:  number,  column:  number,  value:  Object): boolean;
                /**
                 * Moves data from one range to another.
                 * @param {number} fromRow The source row.
                 * @param {number} fromColumn The source column.
                 * @param {number} toRow The target row.
                 * @param {number} toColumn The target column.
                 * @param {number} rowCount The row count.
                 * @param {number} columnCount The column count.
                 * @param {GC.Spread.Sheets.CopyToOptions} option The copy option.
                 */
                moveTo(fromRow:  number,  fromColumn:  number,  toRow:  number,  toColumn:  number,  rowCount:  number,  columnCount:  number,  option:  GC.Spread.Sheets.CopyToOptions): void;
                /**
                 * Gets or sets the name of the worksheet.
                 * @param {string} The name of the worksheet.
                 * @returns {string|GC.Spread.Sheets.Worksheet} If no value is set, returns the worksheet name; otherwise, returns the worksheet.
                 */
                name(value?:  string): any;
                /**
                 * Gets or sets the print information for the sheet.
                 * @param {GC.Spread.Sheets.Print.PrintInfo} value The print information for the sheet.
                 * @returns {GC.Spread.Sheets.Print.PrintInfo | GC.Spread.Sheets.Worksheet} If no value is set, returns the print information for the sheet; otherwise, returns the sheet.
                 */
                printInfo(value?:  GC.Spread.Sheets.Print.PrintInfo): any;
                /**
                 * Recalculates all the formulas in the sheet.
                 * @param {boolean} refreshAll Specifies whether to rebuild all fromula reference, custom name and custom functions.
                 */
                recalcAll(): any;
                /**
                 * Removes a custom function.
                 * @param {string} fnName The custom function name.
                 */
                removeCustomFunction(name:  string): void;
                /**
                 * Removes the specified custom name.
                 * @param {string} fnName The custom name.
                 */
                removeCustomName(name:  string): void;
                /**
                 * Removes a style from the Worksheet named styles collection which has the specified name.
                 * @param {string} name The name of the style to remove.
                 */
                removeNamedStyle(name:  string): void;
                /**
                 * Removes the span that contains a specified anchor cell in the specified sheet area.
                 * @param {number} row The row index of the anchor cell for the span (at which spanned cells start).
                 * @param {number} col The column index of the anchor cell for the span (at which spanned cells start).
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 */
                removeSpan(row:  number,  col:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Removes the sparkline for the specified cell.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 */
                removeSparkline(row:  number,  col:  number): void;
                /**
                 * Repaints the specified rectangle.
                 * @param {GC.Spread.Sheets.Rect} clipRect The rectangle to repaint.
                 */
                repaint(clipRect?:  GC.Spread.Sheets.Rect): void;
                /**
                 * Resets the sheet.
                 */
                reset(): void;
                /**
                 * Resumes the calculation service.
                 * @param {boolean} recalcAll Specifies whether to recalculate all formulas.
                 */
                resumeCalcService(recalcAll?:  boolean): void;
                /**
                 * Resumes the event.
                 */
                resumeEvent(): void;
                /**
                 * Resumes the paint.
                 */
                resumePaint(): void;
                /**
                 * Gets or sets the row filter for the sheet.
                 * @param {GC.Spread.Sheets.Filter.RowFilterBase} value The row filter for the sheet.
                 * @returns {GC.Spread.Sheets.Filter.RowFilterBase} The row filter for the sheet.
                 */
                rowFilter(value?:  GC.Spread.Sheets.Filter.RowFilterBase): GC.Spread.Sheets.Filter.RowFilterBase;
                /**
                 * Searches the specified content.
                 * @param {GC.Spread.Sheets.Search.SearchCondition} searchCondition The search condition.
                 * @returns {GC.Spread.Sheets.Search.SearchResult} The search result.
                 */
                search(searchCondition:  GC.Spread.Sheets.Search.SearchCondition): GC.Spread.Sheets.Search.SearchResult;
                /**
                 * Gets or sets whether users can select ranges of items on a sheet.
                 * @param {GC.Spread.Sheets.SelectionPolicy} value Whether users can select single items, ranges, or a combination of both.
                 * @returns {GC.Spread.Sheets.SelectionPolicy|GC.Spread.Sheets.Worksheet} If no value is set, returns the selection policy setting; otherwise, returns the sheet.
                 */
                selectionPolicy(value?:  GC.Spread.Sheets.SelectionPolicy): any;
                /**
                 * Gets or sets whether users can select cells, rows, or columns on a sheet.
                 * @param {GC.Spread.Sheets.SelectionUnit} value Whether users can select cells, rows, or columns.
                 * @returns {GC.Spread.Sheets.SelectionUnit|GC.Spread.Sheets.Worksheet} If no value is set, returns the selection unit setting; otherwise, returns the sheet.
                 */
                selectionUnit(value?:  GC.Spread.Sheets.SelectionUnit): any;
                /**
                 * Sets the active cell for this sheet.
                 * @param {number} row The row index of the cell.
                 * @param {number} col The column index of the cell.
                 * @param {number} rowViewportIndex The row viewport index of the cell.
                 * @param {number} colViewportIndex The column viewport index of the cell.
                 */
                setActiveCell(row:  number,  col:  number,  rowViewportIndex?:  number,  colViewportIndex?:  number): void;
                /**
                 * Sets the values in the specified two-dimensional array of objects into the specified range of cells on this sheet.
                 * @param {number} row The row index.
                 * @param {number} column The column index.
                 * @param {Array} array The array from which to set values.
                 * @param {boolean} setFormula If <c>true</c>, set formulas; otherwise, set values.
                 */
                setArray(row:  number,  column:  number,  array:  any[],  setFormula?:  boolean): void;
                /**
                 * Sets a formula in a specified cell in the specified sheet area.
                 * @param {number} row The start row index.
                 * @param {number} col The start column index.
                 * @param {number} rowCount The number of rows in range.
                 * @param {number} colCount The number of columns in range.
                 * @param {string} value The array formula to place in the specified range.
                 */
                setArrayFormula(row:  number,  col:  number,  rowCount:  number,  colCount:  number,  value:  string): void;
                /**
                 * Sets the binding path for cell-level binding in a specified cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {string} path The binding path for the cell binding source.
                 */
                setBindingPath(row:  number,  col:  number,  path:  string): void;
                /**
                 * Sets the cell type.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.CellTypes.Base} value The cell type.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 */
                setCellType(row:  number,  col:  number,  value:  GC.Spread.Sheets.CellTypes.Base,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets the column count in the specified sheet area.
                 * @param {number} colCount The column count.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 */
                setColumnCount(colCount:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets whether a forced page break is inserted before the specified column on this sheet when printing.
                 * @param {number} column The column index.
                 * @param {boolean} value Set to <c>true</c> to force a page break before the specified column on this sheet when printing.
                 */
                setColumnPageBreak(column:  number,  value:  boolean): void;
                /**
                 * Sets whether users can resize the specified column in the specified sheet area.
                 * @param {number} col The column index.
                 * @param {boolean} value Set to <c>true</c> to allow users to resize the column.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 */
                setColumnResizable(col:  number,  value:  boolean,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets whether a column in the specified sheet area is displayed.
                 * @param {number} col The column index.
                 * @param {boolean} value Whether to display the column.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 */
                setColumnVisible(col:  number,  value:  boolean,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets the width in pixels for the specified column in the specified sheet area.
                 * @param {number} col The column index.
                 * @param {number} value The width in pixels.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to viewport.
                 */
                setColumnWidth(col:  number,  value:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets delimited text (CSV) in the sheet.
                 * @param {number} row The start row.
                 * @param {number} column The start column.
                 * @param {string} text The delimited text.
                 * @param {string} rowDelimiter The row delimiter.
                 * @param {string} columnDelimiter The column delimiter.
                 */
                setCsv(row:  number,  column:  number,  text:  string,  rowDelimiter:  string,  columnDelimiter:  string): void;
                /**
                 * Sets the data source that populates the sheet.
                 * @param {Object} data The data source.
                 * @param {boolean} reset <c>true</c> if the sheet is reset; otherwise, <c>false</c>.
                 */
                setDataSource(data:  any,  reset?:  boolean): void;
                /**
                 * Sets the cell data validator.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.DataValidation.DefaultDataValidator} value The data validator.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 */
                setDataValidator(row:  number,  col:  number,  value:  GC.Spread.Sheets.DataValidation.DefaultDataValidator,  sheetArea:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets the default style information for the sheet.
                 * @param {GC.Spread.Sheets.Style} style The style to set.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 */
                setDefaultStyle(style:  GC.Spread.Sheets.Style,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets the cell formatter.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {Object} value The formatter string or object.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 */
                setFormatter(row:  number,  col:  number,  value:  Object,  sheetArea:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets a formula in a specified cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {string} value The formula to place in the specified cell.
                 */
                setFormula(row:  number,  col:  number,  value:  string): void;
                /**
                 * Sets the row count in the specified sheet area.
                 * @param {number} rowCount The row count.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 */
                setRowCount(rowCount:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets the height in pixels for the specified row in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} value The height in pixels.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 */
                setRowHeight(row:  number,  value:  number,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets whether a forced page break is inserted before the specified row on this sheet when printing.
                 * @param {number} row The row index.
                 * @param {boolean} value Set to <c>true</c> to force a page break before the specified row on this sheet when printing.
                 */
                setRowPageBreak(row:  number,  value:  boolean): void;
                /**
                 * Sets whether users can resize the specified row in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {boolean} value Set to <c>true</c> to let the users resize the specified row.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 */
                setRowResizable(row:  number,  value:  boolean,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets whether the control displays the specified row in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {boolean} value Set to <c>true</c> to display the specified row.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not given, it defaults to <b>viewport</b>.
                 */
                setRowVisible(row:  number,  value:  boolean,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets the selection to a cell or a range and sets the active cell to the first cell.
                 * @param {number} row The row index of the first cell to add.
                 * @param {number} column The column index of the first cell to add.
                 * @param {number} rowCount The number of rows to add.
                 * @param {number} columnCount The number of columns to add.
                 */
                setSelection(row:  number,  column:  number,  rowCount:  number,  columnCount:  number): void;
                /**
                 * Sets the sparkline for a cell.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.Range} dataRange The data range.
                 * @param {GC.Spread.Sheets.Sparklines.DataOrientation} dataOrientation The data orientation.
                 * @param {GC.Spread.Sheets.Sparklines.SparklineType} sparklineType The sparkline type.
                 * @param {GC.Spread.Sheets.Sparklines.SparklineSetting} sparklineSetting The sparkline setting.
                 * @param {GC.Spread.Sheets.Range} dateAxisRange The date axis range.
                 * @param {GC.Spread.Sheets.Sparklines.DataOrientation} dateAxisOrientation The date axis range orientation.
                 * @returns {GC.Spread.Sheets.Sparklines.Sparkline} The sparkline.
                 */
                setSparkline(row:  number,  col:  number,  dataRange:  GC.Spread.Sheets.Range,  dataOrientation:  GC.Spread.Sheets.Sparklines.DataOrientation,  sparklineType:  GC.Spread.Sheets.Sparklines.SparklineType,  sparklineSetting:  GC.Spread.Sheets.Sparklines.SparklineSetting,  dateAxisRange?:  GC.Spread.Sheets.Range,  dateAxisOrientation?:  GC.Spread.Sheets.Sparklines.DataOrientation): GC.Spread.Sheets.Sparklines.Sparkline;
                /**
                 * Sets the style information for a specified cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} column The column index.
                 * @param {GC.Spread.Sheets.Style} value The cell style.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 */
                setStyle(row:  number,  col:  number,  value:  GC.Spread.Sheets.Style,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets the specified style name for a specified cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} column The column index.
                 * @param {string} value The name of the style to set.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 */
                setStyleName(row:  number,  column:  number,  value:  string,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets the tag value for the specified cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {Object} tag The tag value to set for the specified cell.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 */
                setTag(row:  number,  col:  number,  tag:  any,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets the formatted text in the cell in the specified sheet area.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {string} value The text for the specified cell.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 */
                setText(row:  number,  col:  number,  value:  string,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Sets the value for the specified cell in the specified sheet area.
                 * @param {nubmer} row The row index.
                 * @param {number} col The column index.
                 * @param {Object} value The value to set for the specified cell.
                 * @param {GC.Spread.Sheets.SheetArea} sheetArea The sheet area. If this parameter is not provided, it defaults to <b>viewport</b>.
                 * @param {boolean} ignoreRecalc Whether to ignore recalculation.
                 */
                setValue(row:  number,  col:  number,  value:  any,  sheetArea?:  GC.Spread.Sheets.SheetArea): void;
                /**
                 * Moves the view of a cell to the specified position in the viewport.
                 * @param {number} row The row index.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.VerticalPosition} verticalPosition The vertical position in which to display the cell.
                 * @param {GC.Spread.Sheets.HorizontalPosition} horizontalPosition The horizontal position in which to display the cell.
                 */
                showCell(row:  number,  col:  number,  verticalPosition:  GC.Spread.Sheets.VerticalPosition,  horizontalPosition:  GC.Spread.Sheets.HorizontalPosition): void;
                /**
                 * Moves the view of a column to the specified position in the viewport.
                 * @param {number} col The column index.
                 * @param {GC.Spread.Sheets.HorizontalPosition} horizontalPosition The horizontal position in which to display the column.
                 */
                showColumn(col:  number,  horizontalPosition:  GC.Spread.Sheets.HorizontalPosition): void;
                /**
                 * Gets or sets whether the column outline (range group) is visible.
                 * @param {boolean} value Whether to display the column outline.
                 * @returns {boolean | GC.Spread.Sheets.Worksheet} If no value is set, returns a value that indicates whether the column outline is displayed on this sheet; otherwise, returns the worksheet.
                 */
                showColumnOutline(value?:  boolean): any;
                /**
                 * Moves the view of a row to the specified position in the viewport.
                 * @param {number} row The row index.
                 * @param {GC.Spread.Sheets.VerticalPosition} verticalPosition The vertical position in which to display the row.
                 */
                showRow(row:  number,  verticalPosition:  GC.Spread.Sheets.VerticalPosition): void;
                /**
                 * Gets or sets whether the row outline (range group) is visible.
                 * @param {boolean} value Whether to display the row outline.
                 * @returns {boolean | GC.Spread.Sheets.Worksheet} If no value is set, returns a value that indicates whether the row outline is displayed on this sheet; otherwise, returns the worksheet.
                 */
                showRowOutline(value?:  boolean): any;
                /**
                 * Sorts a range of cells in this sheet in the data model.
                 * @param {number} row The index of the starting row of the block of cells to sort.
                 * @param {number} column The index of the starting column of the block of cells to sort.
                 * @param {number} rowCount The number of rows in the block of cells.
                 * @param {number} columnCount The number of columns in the block of cells.
                 * @param {boolean} byRows Set to <c>true</c> to sort by rows, and <c>false</c> to sort by columns.
                 * @param {Object} sortInfo The SortInfo object with sort criteria and information about how to perform the sort. For example, [{index:0,ascending:true}]
                 * @returns {boolean} <c>true</c> if the data is sorted successfully; otherwise, <c>false</c>.
                 */
                sortRange(row:  number,  column:  number,  rowCount:  number,  columnCount:  number,  byRows:  boolean,  sortInfo:  Object): boolean;
                /**
                 * Starts to edit the cell.
                 * @param {boolean} selectAll Set to <c>true</c> to select all the text in the cell.
                 * @param {string} defaultText The default text to display while editing the cell.
                 */
                startEdit(selectAll?:  boolean,  defaultText?:  string): void;
                /**
                 * Suspends the calculation service.
                 * @param {boolean} ignoreDirty Specifies whether to invalidate the dependency cells.
                 */
                suspendCalcService(ignoreDirty?:  boolean): void;
                /**
                 * Suspends the event.
                 */
                suspendEvent(): void;
                /**
                 * Suspends the paint.
                 */
                suspendPaint(): void;
                /**
                 * Gets or sets the tag value for the current sheet.
                 * @param {Object} value The tag value to set for the current sheet.
                 * @returns {Object|GC.Spread.Sheets.Worksheet} If no value is set, returns the tag value of the current sheet; otherwise, returns the worksheet.
                 */
                tag(value?:  any): any;
                /**
                 * Saves the object state to a JSON string.
                 * @param {Object} serializationOption Serialization option that contains the <i>includeBindingSource</i> argument. See the Remarks for more information.
                 * @returns {Object} The sheet data.
                 */
                toJSON(serializationOption?:  Object): Object;
                /**
                 * Removes the binding of an event to the sheet.
                 * @param {string} type The event type.
                 * @param {Function} fn Specifies the function for which to remove the binding.
                 */
                unbind(type:  any,  fn?:  any): void;
                /**
                 * Removes the binding of all events to the sheet.
                 */
                unbindAll(): void;
                /**
                 * Ungroups the sparklines in the specified group.
                 * @param {GC.Spread.Sheets.Sparklines.SparklineGroup} group The sparkline group.
                 */
                ungroupSparkline(group:  GC.Spread.Sheets.Sparklines.SparklineGroup): void;
                /**
                 * Sets whether the worksheet is displayed.
                 * @param {boolean} value Whether the worksheet is displayed.
                 * @returns {Object} If you call this function without a parameter, it returns a boolean indicating whether the sheet is visible;
                 * otherwise, it returns the current worksheet object.
                 */
                visible(value?:  boolean): any;
                /**
                 * Gets or sets the zoom factor for the sheet.
                 * @param {number} factor The zoom factor.
                 * @returns {number|GC.Spread.Sheets.Worksheet} If no value is set, returns the zoom factor; otherwise, returns the worksheet.
                 */
                zoom(factor?:  number): any;
            }
            module Bindings{

                export class CellBindingSource{
                    /**
                     * Represents a source for cell binding.
                     * @param {Object} source The data source.
                     * @class
                     */
                    constructor(source:  Object);
                    /**
                     * Gets the wrapped data source for cell binding.
                     * @returns {Object} The original data source.
                     */
                    getSource(): Object;
                    /**
                     * Gets the value of the source by the binding path.
                     * @param {string} path The binding path.
                     * @returns {Object} Returns the value of the binding source at the specified path.
                     */
                    getValue(path:  string): Object;
                    /**
                     * Sets the value of the source by the binding path.
                     * @param {string} path The row index.
                     * @param {Object} value The value to set.
                     */
                    setValue(path:  string,  value:  Object): void;
                }
            }

            module CalcEngine{
                /**
                 * Evaluates the specified formula.
                 * @param {object} context The evaluation context; in general, you should use the active sheet object.
                 * @param {string} formula The formula string.
                 * @param {number} baseRow The base row index of the formula.
                 * @param {number} baseColumn The base column index of the formula.
                 * @param {boolean} useR1C1 Whether to use the r1c1 reference style.
                 * @returns {object} The evaluated formula result.
                 */
                function evaluateFormula(context:  Object,  formula:  string,  baseRow?:  number,  baseColumn?:  number,  useR1C1?:  boolean): any;
                /**
                 * Unparse the specified expression tree to formula string.
                 * @param {object} context The context; in general, you should use the active sheet object.
                 * @param {GC.Spread.CalcEngine.Expression} expression The expression tree.
                 * @param {number} baseRow The base row index of the formula.
                 * @param {number} baseColumn The base column index of the formula.
                 * @param {boolean} useR1C1 Whether to use the r1c1 reference style.
                 * @returns {string} The formula string.
                 */
                function expressionToFormula(context:  Object,  expression:  GC.Spread.CalcEngine.Expression,  baseRow?:  number,  baseColumn?:  number,  useR1C1?:  boolean): string;
                /**
                 * Parse the specified formula to expression tree.
                 * @param {object} context The context; in general, you should use the active sheet object.
                 * @param {string} formula The formula string.
                 * @param {number} baseRow The base row index of the formula.
                 * @param {number} baseColumn The base column index of the formula.
                 * @param {boolean} useR1C1 Whether to use the r1c1 reference style.
                 * @returns {GC.Spread.CalcEngine.Expression} The expression tree.
                 */
                function formulaToExpression(context:  Object,  formula:  string,  baseRow?:  number,  baseColumn?:  number,  useR1C1?:  boolean): GC.Spread.CalcEngine.Expression;
                /**
                 * Converts the specified cell range to a formula string.
                 * @param {GC.Spread.Sheets.Worksheet} sheet The base sheet.
                 * @param {string} formula The formula.
                 * @param {number} baseRow The base row index of the formula.
                 * @param {number} baseCol The base column index of the formula.
                 * @returns {Array} The cell ranges that refers to the formula string.
                 */
                function formulaToRanges(sheet: any,  formula: any,  baseRow?:  number,  baseCol?:  number): Object[];
                /**
                 * Converts the specified cell range to a formula string.
                 * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell range in the sheet.
                 * @param {number} baseRow The base row index of the formula.
                 * @param {number} baseCol The base column index of the formula.
                 * @param {GC.Spread.CalcEngine.RangeReferenceRelative} rangeReferenceRelative Whether the range reference is relative or absolute.
                 * @param {boolean} useR1C1 Whether to use the r1c1 reference style.
                 * @returns {string} The formula string that refers to the specified cell range.
                 */
                function rangesToFormula(ranges:  Range[],  baseRow?:  number,  baseCol?:  number,  rangeReferenceRelative?:  RangeReferenceRelative,  useR1C1?:  boolean): string;
                /**
                 * Converts the specified cell range to a formula string.
                 * @param {GC.Spread.Sheets.Range} range The cell range in the sheet.
                 * @param {number} baseRow The base row index of the formula.
                 * @param {number} baseCol The base column index of the formula.
                 * @param {GC.Spread.CalcEngine.RangeReferenceRelative} rangeReferenceRelative Whether the range reference is relative or absolute.
                 * @param {boolean} useR1C1 Whether to use the r1c1 reference style.
                 * @returns {string} The formula string that refers to the specified cell range.
                 */
                function rangeToFormula(range:  Range,  baseRow?:  number,  baseCol?:  number,  rangeReferenceRelative?:  RangeReferenceRelative,  useR1C1?:  boolean): string;
                /**
                 * Specifies whether the range reference is relative or absolute.
                 * @enum {number}
                 */
                export enum RangeReferenceRelative{
                    /**
                     * Specifies all reference is absolute
                     */
                    allAbsolute= 0,
                    /**
                     * Specifies start row is relative
                     */
                    startRowRelative= 1,
                    /**
                     * Specifies start column is relative
                     */
                    startColRelative= 2,
                    /**
                     * Specifies end row is relative
                     */
                    endRowRelative= 4,
                    /**
                     * Specifies end column is relative
                     */
                    endColRelative= 8,
                    /**
                     * Specifies row is relative
                     */
                    rowRelative= 5,
                    /**
                     * Specifies column is relative
                     */
                    colRelative= 10,
                    /**
                     * Specifies all reference is relative
                     */
                    allRelative= 15
                }

            }

            module CellTypes{
                /**
                 * Specifies the text alignment for check box cells.
                 * @enum {number}
                 */
                export enum CheckBoxTextAlign{
                    /**
                     * Specifies text is on top of the check box.
                     */
                    top= 0,
                    /**
                     * Specifies text is below the check box.
                     */
                    bottom= 1,
                    /**
                     * Specifies text is to the left of the check box.
                     */
                    left= 2,
                    /**
                     * Specifies text is to the right of the check box.
                     */
                    right= 3
                }

                /**
                 *  Specifies what is written out to the data model for a selected item from
                 *  certain cell types that offer a selection of multiple values.
                 * @readonly
                 * @enum {number}
                 */
                export enum EditorValueType{
                    /**
                     *  Writes the text value of the selected item to the model.
                     */
                    text= 0,
                    /**
                     * Writes the index of the selected item to the model.
                     */
                    index= 1,
                    /**
                     *  Writes the corresponding data value of the selected item to the model.
                     */
                    value= 2
                }

                /**
                 * Specifies the hyperlink's target type.
                 * @enum {number}
                 */
                export enum HyperLinkTargetType{
                    /**
                     * Opens the hyperlinked document in a new window or tab.
                     */
                    blank= 0,
                    /**
                     * Opens the hyperlinked document in the same frame where the user clicked.
                     */
                    self= 1,
                    /**
                     * Opens the hyperlinked document in the parent frame.
                     */
                    parent= 2,
                    /**
                     * Opens the hyperlinked document in the full body of the window.
                     */
                    top= 3
                }


                export class Base{
                    /**
                     * Represents the base class for the other cell type classes.
                     * @class
                     */
                    constructor();
                    /**
                     * Represents the type name string used for supporting serialization.
                     * @type {string}
                     */
                    typeName: string;
                    /**
                     * Activates the editor, including setting properties or attributes for the editor and binding events for the editor.
                     * @param {object} editorContext The DOM element that was created by the createEditorElement method.
                     * @param {GC.Spread.Sheets.Style} cellStyle The cell's actual style.
                     * @param {GC.Spread.Sheets.Rect} cellRect The cell's layout information.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     */
                    activateEditor(editorContext:  any,  cellStyle:  GC.Spread.Sheets.Style,  cellRect:  GC.Spread.Sheets.Rect,  context?:  any): void;
                    /**
                     * Creates a DOM element then returns it.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     * @returns {object} Returns a DOM element.
                     */
                    createEditorElement(context?:  any): any;
                    /**
                     * Deactivates the editor, such as unbinding events for editor.
                     * @param {object} editorContext The DOM element that was created by the createEditorElement method.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     */
                    deactivateEditor(editorContext:  any,  context?:  any): void;
                    /**
                     * Focuses the editor DOM element.
                     * @param {object} editorContext The DOM element that was created by the createEditorElement method.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     */
                    focus(editorContext:  any,  context?:  any): void;
                    /**
                     * Formats a value with the specified format to a string.
                     * @param {object} value The object value to format.
                     * @param {GC.Spread.Formatter.GeneralFormatter} format The format.
                     * @param {object} conditionalForeColor The conditional foreground color.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     * @returns {string} Returns the formatted string.
                     */
                    format(value:  any,  format:  any,  conditionalForeColor?:  any,  context?:  any): string;
                    /**
                     * Loads the object state from the specified JSON string.
                     * @param {Object} settings The cell type data from deserialization.
                     */
                    fromJSON(settings:  Object): void;
                    /**
                     * Gets a cell's height that can be used to handle the row's automatic fit.
                     * @param {object} value The cell's value.
                     * @param {string} text The cell's text.
                     * @param {GC.Spread.Sheets.Style} cellStyle The cell's actual value.
                     * @param {number} zoomFactor The current sheet's zoom factor.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     * @returns {number} Returns the cell's height that can be used to handle the row's automatic fit.
                     */
                    getAutoFitHeight(value:  any,  text:  string,  cellStyle:  GC.Spread.Sheets.Style,  zoomFactor:  number,  context?:  any): number;
                    /**
                     * Gets a cell's width that can be used to handle the column's automatic fit.
                     * @param {object} value The cell's value.
                     * @param {string} text The cell's text.
                     * @param {GC.Spread.Sheets.Style} cellStyle The cell's actual value.
                     * @param {number} zoomFactor The current sheet's zoom factor.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     * @returns {number} Returns the cell's width that can be used to handle the column's automatic fit.
                     */
                    getAutoFitWidth(value:  any,  text:  string,  cellStyle:  GC.Spread.Sheets.Style,  zoomFactor:  number,  context?:  any): number;
                    /**
                     * Gets the editor's value.
                     * @param {object} editorContext The DOM element that was created by the createEditorElement method.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     * @returns {object} Returns the editor's value.
                     */
                    getEditorValue(editorContext:  any,  context?:  any): any;
                    /**
                     * Gets the cell type's hit information.
                     * @param {number} x <i>x</i>-coordinate of pointer's current location relative to the canvas.
                     * @param {number} y <i>y</i>-coordinate of pointer's current location relative to the canvas.
                     * @param {GC.Spread.Sheets.Style} cellStyle The current cell's actual style.
                     * @param {GC.Spread.Sheets.Rect} cellRect The current cell's layout information.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     * @returns {object} Returns an object that contains the <i>x</i>, <i>y</i>, <i>row</i>, <i>col</i>, <i>cellRect</i>, and <i>sheetArea</i> parameters, and a value to indicate <i>isReservedLocation</i>.
                     * <i>isReservedLocation</i> is <c>true</c> if the hit test is in a special area that the cell type needs to handle; otherwise, <c>false</c>.
                     */
                    getHitInfo(x:  number,  y:  number,  cellStyle:  GC.Spread.Sheets.Style,  cellRect:  GC.Spread.Sheets.Rect,  context?:  any): GC.Spread.Sheets.IHitTestCellTypeHitInfo;
                    /**
                     * Whether the editing value has changed.
                     * @param {object} oldValue Old editing value.
                     * @param {object} newValue New editing value.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     * @returns {boolean} <c>true</c> if oldValue equals newValue; otherwise, <c>false</c>.
                     */
                    isEditingValueChanged(oldValue:  any,  newValue:  any,  context?:  any): boolean;
                    /**
                     * Whether this cell type is aware of IME.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     * @returns {boolean} <c>true</c> if the cell type is aware of IME; otherwise, <c>false</c>.
                     */
                    isImeAware(context?:  any): boolean;
                    /**
                     * Whether the cell type handles the keyboard event itself.
                     * @param {KeyboardEvent} e The KeyboardEvent.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     * @returns {boolean} Returns <c>true</c> if the cell type handles the keyboard event itself; otherwise, <c>false</c>.
                     */
                    isReservedKey(e:  any,  context?:  any): boolean;
                    /**
                     * Paints a cell on the canvas.
                     * @param {CanvasRenderingContext2D} ctx The canvas's two-dimensional context.
                     * @param {object} value The cell's value.
                     * @param {number} x <i>x</i>-coordinate relative to the canvas.
                     * @param {number} y <i>y</i>-coordinate relative to the canvas.
                     * @param {number} w The cell's width.
                     * @param {number} h The cell's height.
                     * @param {GC.Spread.Sheets.Style} style The cell's actual style.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     */
                    paint(ctx:  CanvasRenderingContext2D,  value:  any,  x:  number,  y:  number,  w:  number,  h:  number,  style:  GC.Spread.Sheets.Style,  context?:  any): void;
                    /**
                     * Paints the cell content area on the canvas.
                     * @param {CanvasRenderingContext2D} ctx The canvas's two-dimensional context.
                     * @param {object} value The cell's value.
                     * @param {number} x <i>x</i>-coordinate relative to the canvas.
                     * @param {number} y <i>y</i>-coordinate relative to the canvas.
                     * @param {number} w The cell content area's width.
                     * @param {number} h The cell content area's height.
                     * @param {GC.Spread.Sheets.Style} style The cell's actual style.
                     * @param {object} context The context associated with the cell type.
                     */
                    paintContent(ctx:  CanvasRenderingContext2D,  value:  any,  x:  number,  y:  number,  w:  number,  h:  number,  style:  GC.Spread.Sheets.Style,  context?:  any): void;
                    /**
                     * Parses the text with the specified format string to an object.
                     * @param {string} text The parse text string.
                     * @param {object} formatStr The parse format string.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     * @returns {object} The parsed object.
                     */
                    parse(text:  string,  formatStr:  string,  context?:  any): any;
                    /**
                     * Processes key down in display mode.
                     * @param {KeyboardEvent} event The KeyboardEvent.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     * @returns {boolean} Returns <c>true</c> if the process is successful; otherwise, <c>false</c>.
                     */
                    processKeyDown(event:  KeyboardEvent,  context?:  any): boolean;
                    /**
                     * Processes key up in display mode.
                     * @param {KeyboardEvent} event The KeyboardEvent.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     * @returns {boolean} Returns <c>true</c> if the process is successful; otherwise, <c>false</c>.
                     */
                    processKeyUp(event:  KeyboardEvent,  context?:  any): boolean;
                    /**
                     * Processes mouse down in display mode.
                     * @param {object} hitInfo The hit test information returned by the getHitInfo method. See the Remarks for more information.
                     * @returns {boolean} Returns <c>true</c> if the process is successful; otherwise, <c>false</c>.
                     */
                    processMouseDown(hitInfo:  GC.Spread.Sheets.IHitTestCellTypeHitInfo): boolean;
                    /**
                     * Processes mouse enter in display mode.
                     * @param {object} hitInfo The hit test information returned by the getHitInfo method. See the Remarks for more information.
                     * @returns {boolean} Returns <c>true</c> if the process is successful; otherwise, <c>false</c>.
                     */
                    processMouseEnter(hitInfo:  GC.Spread.Sheets.IHitTestCellTypeHitInfo): boolean;
                    /**
                     * Processes mouse leave in display mode.
                     * @param {object} hitInfo The hit test information returned by the getHitInfo method. See the Remarks for more information.
                     * @returns {boolean} Returns <c>true</c> if the process is successful; otherwise, <c>false</c>.
                     */
                    processMouseLeave(hitInfo:  GC.Spread.Sheets.IHitTestCellTypeHitInfo): boolean;
                    /**
                     * Processes mouse move in display mode.
                     * @param {object} hitInfo The hit test information returned by the getHitInfo method. See the Remarks for more information.
                     * @returns {boolean} Returns <c>true</c> if the process is successful; otherwise, <c>false</c>.
                     */
                    processMouseMove(hitInfo:  GC.Spread.Sheets.IHitTestCellTypeHitInfo): boolean;
                    /**
                     * Processes mouse up in display mode.
                     * @param {object} hitInfo The hit test information returned by the getHitInfo method. See the Remarks for more information.
                     * @returns {boolean} Returns <c>true</c> if the process is successful; otherwise, <c>false</c>.
                     */
                    processMouseUp(hitInfo:  GC.Spread.Sheets.IHitTestCellTypeHitInfo): boolean;
                    /**
                     * Selects all the text in the editor DOM element.
                     * @param {object} editorContext The DOM element that was created by the createEditorElement method.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     */
                    selectAll(editorContext:  any,  context?:  any): void;
                    /**
                     * Sets the editor's value.
                     * @param {object} editorContext The DOM element that was created by the createEditorElement method.
                     * @param {object} value The value returned from the active cell.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     */
                    setEditorValue(editorContext:  any,  value:  any,  context?:  any): void;
                    /**
                     * Saves the object state to a JSON string.
                     * @returns {Object} The cell type data.
                     */
                    toJSON(): Object;
                    /**
                     * Updates the editor's size.
                     * @param {object} editorContext The DOM element that was created by the createEditorElement method.
                     * @param {GC.Spread.Sheets.Style} cellStyle The cell's actual style.
                     * @param {GC.Spread.Sheets.Rect} cellRect The cell's layout information.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     * @returns {object} Returns the new size for cell wrapper element, it should contain two properties 'width' and 'height'.
                     */
                    updateEditor(editorContext:  any,  cellStyle:  GC.Spread.Sheets.Style,  cellRect:  GC.Spread.Sheets.Rect,  context?:  any): any;
                    /**
                     * Updates the cell wrapper element size.
                     * @param {object} editorContext The DOM element that was created by the createEditorElement method.
                     * @param {object} editorBounds The cell wrapper element's new size.<br />
                     ** editorBounds.width {number} The cell wrapper element's new width value.<br />
                     ** editorBounds.height {number} The cell wrapper element's new height value.
                     * @param {GC.Spread.Sheets.Style} cellStyle The cell's actual style.
                     */
                    updateEditorContainer(editorContext:  any,  editorBounds:  Object,  cellStyle:  GC.Spread.Sheets.Style): void;
                    /**
                     * Updates the editor's ime-mode.
                     * @param {object} editorContext The DOM element that was created by the createEditorElement method.
                     * @param {GC.Spread.Sheets.ImeMode} imeMode The ime-mode from cell's actual style.
                     * @param {object} context The context associated with the cell type. See the Remarks for more information.
                     */
                    updateImeMode(editorContext:  any,  imeMode:  GC.Spread.Sheets.ImeMode,  context?:  any): void;
                }

                export class Button extends Base{
                    /**
                     * Represents a button cell.
                     * @extends GC.Spread.Sheets.CellTypes.Base
                     * @class
                     */
                    constructor();
                    /**
                     * Gets or sets the button's background color.
                     * @param {string} value The button's background color.
                     * @returns {string | GC.Spread.Sheets.CellTypes.Button} If no value is set, returns the background color; otherwise, returns the button cell type.
                     */
                    buttonBackColor(value?:  string): any;
                    /**
                     * Gets or sets the button's bottom margin in pixels relative to the cell.
                     * @param {number} value The button's bottom margin relative to the cell.
                     * @returns {number | GC.Spread.Sheets.CellTypes.Button} If no value is set, returns the bottom margin in pixels; otherwise, returns the button cell type.
                     */
                    marginBottom(value?:  number): any;
                    /**
                     * Gets or sets the button's left margin in pixels relative to the cell.
                     * @param {number} value The button's left margin relative to the cell.
                     * @returns {number | GC.Spread.Sheets.CellTypes.Button} If no value is set, returns the left margin in pixels; otherwise, returns the button cell type.
                     */
                    marginLeft(value?:  number): any;
                    /**
                     * Gets or sets the button's right margin in pixels relative to the cell.
                     * @param {number} value The button's right margin relative to the cell.
                     * @returns {number | GC.Spread.Sheets.CellTypes.Button} If no value is set, returns the right margin in pixels; otherwise, returns the button cell type.
                     */
                    marginRight(value?:  number): any;
                    /**
                     * Gets or sets the button's top margin in pixels relative to the cell.
                     * @param {number} value The button's top margin relative to the cell.
                     * @returns {number | GC.Spread.Sheets.CellTypes.Button} If no value is set, returns the top margin in pixels; otherwise, returns the button cell type.
                     */
                    marginTop(value?:  number): any;
                    /**
                     * Gets or sets the button's content.
                     * @param {string} value The button's content.
                     * @returns {string | GC.Spread.Sheets.CellTypes.Button} If no value is set, returns the content; otherwise, returns the button cell type.
                     */
                    text(value?:  string): any;
                }

                export class CheckBox extends Base{
                    /**
                     * Represents a check box cell.
                     * @extends GC.Spread.Sheets.CellTypes.Base
                     * @class
                     */
                    constructor();
                    /**
                     * Gets or sets the caption of the cell type.
                     * @param {string} value The caption of the cell type.
                     * @returns {string | GC.Spread.Sheets.CellTypes.CheckBox} If no value is set, returns the caption; otherwise, returns the check box cell type.
                     */
                    caption(value?:  string): any;
                    /**
                     * Gets or sets a value that indicates whether the check box supports three states.
                     * @param {boolean} value Whether the check box supports three states.
                     * @returns {boolean | GC.Spread.Sheets.CellTypes.CheckBox} If no value is set, returns whether the check box supports three states; otherwise, returns the check box cell type.
                     */
                    isThreeState(value?:  boolean): any;
                    /**
                     * Gets or sets the text alignment relative to the check box.
                     * @param {GC.Spread.Sheets.CellTypes.CheckBoxTextAlign} value The text alignment relative to the check box.
                     * @returns {GC.Spread.Sheets.CellTypes.CheckBoxTextAlign | GC.Spread.Sheets.CellTypes.CheckBox} If no value is set, returns the text alignment relative to the check box; otherwise, returns the check box cell type.
                     */
                    textAlign(value?:  GC.Spread.Sheets.CellTypes.CheckBoxTextAlign): any;
                    /**
                     * Gets or sets the text in the cell when the cell's value is <c>false</c>.
                     * @param {string} value The text in the cell when the cell's value is <c>false</c>.
                     * @returns {string | GC.Spread.Sheets.CellTypes.CheckBox} If no value is set, returns the text in the cell when the cell's value is <c>false</c>. If a value is set, returns the check box cell type.
                     */
                    textFalse(value?:  string): any;
                    /**
                     * Gets or sets the text in the cell when the cell's value is indeterminate (neither <c>true</c> nor <c>false</c>).
                     * @param {string} value The text in the cell when the cell's value is indeterminate.
                     * @returns {string | GC.Spread.Sheets.CellTypes.CheckBox} If no value is set, returns the text in the cell when the cell's value is indeterminate. If a value is set, returns the check box cell type.
                     */
                    textIndeterminate(value?:  string): any;
                    /**
                     * Gets or sets the text in the cell when the cell's value is <c>true</c>.
                     * @param {string} value The text when the cell's value is <c>true</c>.
                     * @returns {string | GC.Spread.Sheets.CellTypes.CheckBox} If no value is set, returns the text when the cell's value is <c>true</c>. If a value is set, returns the check box cell type.
                     */
                    textTrue(value?:  string): any;
                }

                export class ColumnHeader extends Base{
                    /**
                     * Represents the painter of the column header cells.
                     * @extends GC.Spread.Sheets.CellTypes.Base
                     * @class
                     */
                    constructor();
                }

                export class ComboBox extends Base{
                    /**
                     * Represents an editable combo box cell.
                     * @extends GC.Spread.Sheets.CellTypes.Base
                     * @class
                     */
                    constructor();
                    /**
                     * Gets or sets whether the combo box is editable.
                     * @param {boolean} value Whether the combo box is editable.
                     * @returns {boolean | GC.Spread.Sheets.CellTypes.ComboBoxBox} If no value is set, returns whether the combo box is editable; otherwise, returns the combo box cellType.
                     */
                    editable(value?:  boolean): any;
                    /**
                     * Gets or sets the value that is written to the underlying data model.
                     * @param {GC.Spread.Sheets.CellTypes.EditorValueType} value The type of editor value.
                     * @returns {GC.Spread.Sheets.CellTypes.EditorValueType | GC.Spread.Sheets.CellTypes.ComboBoxBox} If no value is set, returns the type of editor value; otherwise, returns the combo box cellType.
                     */
                    editorValueType(value?:  GC.Spread.Sheets.CellTypes.EditorValueType): any;
                    /**
                     * Gets or sets the height of each item.
                     * @param {number} value The height of each item.
                     * @returns {number | GC.Spread.Sheets.CellTypes.ComboBoxBox} If no value is set, returns the height of each item; otherwise, returns the  combo box cellType.
                     */
                    itemHeight(value?:  number): any;
                    /**
                     * Gets or sets the items for the drop-down list in the combo box.
                     * @param {Array} items The items for the combo box.
                     * @returns {Array | GC.Spread.Sheets.CellTypes.ComboBoxBox} If no value is set, returns the items array; otherwise, returns the  combo box cellType.
                     */
                    items(items?:  any[]): any;
                    /**
                     * Gets or sets the maximum item count of the drop-down list per page.
                     * @param {number} value The maximum item count of the drop-down list per page.
                     * @returns {number | GC.Spread.Sheets.CellTypes.ComboBoxBox} If no value is set, returns the maximum item count of the drop-down list per page; otherwise, returns the  combo box cellType.
                     */
                    maxDropDownItems(value?:  number): any;
                }

                export class Corner extends Base{
                    /**
                     * Represents the painter of the corner cell.
                     * @extends GC.Spread.Sheets.CellTypes.Base
                     * @class
                     */
                    constructor();
                }

                export class HyperLink extends Base{
                    /**
                     * Represents the hyperlink cell.
                     * @extends GC.Spread.Sheets.CellTypes.Base
                     * @class
                     */
                    constructor();
                    /**
                     * Gets or sets the color of the hyperlink.
                     * @param {string} value The hyperlink color.
                     * @returns {string | GC.Spread.Sheets.CellTypes.HyperLink} If no value is set, returns the hyperlink color; otherwise, returns the hyperLink cell type.
                     */
                    linkColor(value?:  string): any;
                    /**
                     * Gets or sets the tooltip for the hyperlink.
                     * @param {string} value The tooltip text.
                     * @returns {string | GC.Spread.Sheets.CellTypes.HyperLink} If no value is set, returns the tooltip text; otherwise, returns the hyperLink cell type.
                     */
                    linkToolTip(value?:  string): any;
                    /**
                     *  Gets or sets the type for the hyperlink's target.
                     * @param {GC.Spread.Sheets.CellTypes.HyperLinkTargetType} value The hyperlink's target type.
                     * @returns {GC.Spread.Sheets.CellTypes.HyperLinkTargetType | GC.Spread.Sheets.CellTypes.HyperLink} If no value is set, returns the hyperlink's target type; otherwise, returns the hyperLink cell type.
                     */
                    target(value?:  GC.Spread.Sheets.CellTypes.HyperLinkTargetType): any;
                    /**
                     * Gets or sets the text string for the hyperlink.
                     * @param {string} value The text displayed in the hyperlink.
                     * @returns {string | GC.Spread.Sheets.CellTypes.HyperLink} If no value is set, returns the text in the hyperlink; otherwise, returns the hyperLink cell type.
                     */
                    text(value?:  string): any;
                    /**
                     * Gets or sets the color of visited links.
                     * @param {string} value The visited link color.
                     * @returns {string | GC.Spread.Sheets.CellTypes.HyperLink} If no value is set, returns the visited link color; otherwise, returns the hyperLink cell type.
                     */
                    visitedLinkColor(value?:  string): any;
                }

                export class RowHeader extends Base{
                    /**
                     * Represents the painter of the row header cells.
                     * @extends GC.Spread.Sheets.CellTypes.Base
                     * @class
                     */
                    constructor();
                }

                export class Text extends Base{
                    /**
                     * Represents a text cell type.
                     * @extends GC.Spread.Sheets.CellTypes.Base
                     * @class
                     */
                    constructor();
                }
            }

            module Commands{
                /**
                 * Represents the command used to automatically resize the column in a sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.columns {Array} The resize columns; each item is an object which has a col.<br />
                 ** options.rowHeader {boolean} Whether the resized columns are in the row header area.<br />
                 ** options.autoFitType {GC.Spread.Sheets.AutoFitType} Whether the auto-fit action includes the header text.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var autoFitColumn: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, columns: Object[], rowHeader: boolean, autoFitType: GC.Spread.Sheets.AutoFitType}, isUndo: boolean): any};
                /**
                 * Represents the command used to automatically resize the row in a sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.rows {Array} The resize rows; each item is an object which has a row.<br />
                 ** options.columnHeader {boolean} Whether the resized rows are in the column header area.<br />
                 ** options.autoFitType {GC.Spread.Sheets.AutoFitType} Whether the auto-fit action includes the header text.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var autoFitRow: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, rows: Object[], columnHeader: boolean, autoFitType: GC.Spread.Sheets.AutoFitType}, isUndo: boolean): any};
                /**
                 * Represents the command used to stop cell editing and cancel input.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var cancelInput: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to switch the formula reference between relative, absolute, and mixed when editing formulas.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var changeFormulaReference: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to clear the cell value.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var clear: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to clear the active cell value and enters edit mode.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var clearAndEditing: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to clear cell values on a worksheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.ranges {Array} The clear cell value ranges whose item type is GC.Spread.Sheets.Range.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var clearValues: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, ranges:GC.Spread.Sheets.Range[]}, isUndo: boolean): any};
                /**
                 * Represents the command used for a Clipboard paste on the worksheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.fromSheet {GC.Spread.Sheets.Worksheet} The source sheet.<br />
                 ** options.fromRanges {Array} The source range array which item type is GC.Spread.Sheets.Range.<br />
                 ** options.pastedRanges {Array} The target range array which item type is GC.Spread.Sheets.Range.<br />
                 ** options.isCutting {boolean} Whether the operation is cutting or copying.<br />
                 ** options.clipboardText {string} The text on the clipboard.<br />
                 ** options.pasteOption {GC.Spread.Sheets.ClipboardPasteOptions} The Clipboard pasting option that indicates which content to paste.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var clipboardPaste: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, fromSheet: GC.Spread.Sheets.Worksheet, fromRanges: GC.Spread.Sheets.Range[], pastedRanges: GC.Spread.Sheets.Range[], isCutting: boolean, clipboardText: string, pasteOption: GC.Spread.Sheets.ClipboardPasteOptions}, isUndo: boolean): any};
                /**
                 * Represents the command used to commit the cell editing and sets the array formula to the active range.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var commitArrayFormula: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to stop cell editing and moves the active cell to the next row.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var commitInputNavigationDown: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to stop cell editing and moves the active cell to the previous row.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var commitInputNavigationUp: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to copy the selected item text to the Clipboard.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var copy: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to cut the selected item text to the Clipboard.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var cut: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command for deleting the floating objects.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and a method <b>execute</b> that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var deleteFloatingObjects: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to drag and copy the floating objects on the sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.floatingObjects {Array} The names array of floating objects.<br />
                 ** options.offsetX The horizontal offset.<br />
                 ** options.offsetY The vertical offset.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise <c>false</c>.
                 */
                var dragCopyFloatingObjects: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, floatingObjects: string[], offsetX: number, offsetY: number}, isUndo: boolean): any};
                /**
                 * Represents the command used to drag a range and drop it onto another range on the worksheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operation,
                 * and a method <b>execute</b> that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * commandOptions {Object} The options of the operation.<br />
                 ** commandOptions.sheetName {string} The sheet name.<br />
                 ** commandOptions.fromRow {number} The source row index for the drag drop.<br />
                 ** commandOptions.fromColumn {number} The source column index for the drag drop.<br />
                 ** commandOptions.toRow {number} The destination row index for the drag drop.<br />
                 ** commandOptions.toColumn {number} The destination column index for the drag drop.<br />
                 ** commandOptions.rowCount {number} The row count for the drag drop.<br />
                 ** commandOptions.columnCount {number} The column count for the drag drop.<br />
                 ** commandOptions.copy {boolean} If set to <c>true</c> copy; otherwise, cut if <c>false</c>.<br />
                 ** commandOptions.insert {boolean} If set to <c>true</c> inserts the drag data in the drop row or column.<br />
                 ** commandOptions.option {GC.Spread.Sheets.CopyToOptions} Indicates the content type to drag and drop.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var dragDrop: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, fromRow: number, fromColumn: number, toRow: number, toColumn: number, rowCount: number, columnCount: number, copy: boolean, insert: boolean, option: GC.Spread.Sheets.CopyToOptions}, isUndo: boolean): any};
                /**
                 * Represents the command used to apply a new value to a cell on the worksheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.row {number} The row index of the cell.<br />
                 ** options.col {number} The column idnex of the cell.<br />
                 ** options.newValue {Object} The new value of the cell.<br />
                 ** options.autoFormat {boolean} Whether to format the new value automatically.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var editCell: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, row: number, col: number, newValue: any, autoFormat: boolean}, isUndo: boolean): any};
                /**
                 * Represents the command to expand or collapse a column range group.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.index {number} The outline summary index.<br />
                 ** options.level {number} The outline level.<br />
                 ** options.collapsed {boolean} Whether to make the outline collapsed or expanded.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var expandColumnOutline: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, index: number, level: number, collapsed: boolean}, isUndo: boolean): any};
                /**
                 * Represents the command used to expand or collapse column range groups on the same level.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.level {number} The outline level.<br />
                 * isUndo {boolean} <c>true</c> if this an undo operation; otherwise, <c>false</c>.
                 */
                var expandColumnOutlineForLevel: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, level: number}, isUndo: boolean): any};
                /**
                 * Represents the command to expand or collapse a row range group.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.index {number} The outline summary index.<br />
                 ** options.level {number} The outline level.<br />
                 ** options.collapsed {boolean} Whether to make the outline collapsed or expanded.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var expandRowOutline: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, index: number, level: number, collapsed: boolean}, isUndo: boolean): any};
                /**
                 * Represents the command used to expand or collapse row range groups on the same level.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.level {number} The outline level.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var expandRowOutlineForLevel: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, level: number}, isUndo: boolean): any};
                /**
                 * Represents the command used to drag and fill a range on the sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operation,
                 * and a method <b>execute</b> that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.startRange {GC.Spread.Sheets.Range} The start range.<br />
                 ** options.fillRange {GC.Spread.Sheets.Range} The fill range.<br />
                 ** options.autoFillType {GC.Spread.Sheets.Fill.AutoFillType} The auto fill type.<br />
                 ** options.fillDirection {GC.Spread.Sheets.Fill.FillDirection} The fill direction.<br />
                 * isUndo {boolean} <c>true</c> if an undo operation; otherwise, <c>false</c>.
                 */
                var fill: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, startRange: GC.Spread.Sheets.Range, fillRange: GC.Spread.Sheets.Range, autoFillType: GC.Spread.Sheets.Fill.AutoFillType, fillDirection: GC.Spread.Sheets.Fill.FillDirection}, isUndo: boolean): any};
                /**
                 * Represents the command for moving floating objects.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.floatingObjects {Array} The names array of floating objects.<br />
                 ** options.offsetX The horizontal offset.<br />
                 ** options.offsetY The vertical offset.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var moveFloatingObjects: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, floatingObjects: string[], offsetX: number, offsetY: number}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the next cell.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var moveToNextCell: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to select the next control if the active cell is the last visible cell; otherwise, move the active cell to the next cell.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var moveToNextCellThenControl: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the previous cell.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var moveToPreviousCell: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to select the previous control if the active cell is the first visible cell; otherwise, move the active cell to the previous cell.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var moveToPreviousCellThenControl: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the last row.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationBottom: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the next row.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationDown: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the last column.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationEnd: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the last column without regard to frozen trailing columns.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationEnd2: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the first cell in the sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationFirst: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the first column.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and a method <b>execute</b> that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationHome: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the first column without regard to frozen columns.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> metod that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationHome2: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the last cell in the sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationLast: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the previous column.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationLeft: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active sheet to the next sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationNextSheet: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell down one page of rows.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationPageDown: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell up one page of rows.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationPageUp: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active sheet to the previous sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationPreviousSheet: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the next column.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationRight: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the first row.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationTop: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to move the active cell to the previous row.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var navigationUp: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command for grouping a column outline (range group) on a sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.index {number} The outline starting index.<br />
                 ** options.count {number} The number of rows or columns to group or ungroup in the outline.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var outlineColumn: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, index: number, count: number}, isUndo: boolean): any};
                /**
                 * Represents the command for grouping a row outline (range group) on a sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.index {number} The outline starting index.<br />
                 ** options.count {number} The number of rows or columns to group or ungroup in the outline.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var outlineRow: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, index: number, count: number}, isUndo: boolean): any};
                /**
                 * Represents the command used to paste the selected items from the Clipboard to the current sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var paste: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command for pasting the floating objects on the sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and a method <b>execute</b> that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var pasteFloatingObjects: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to perform a redo of the most recently undone edit or action.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var redo: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command for ungrouping a column outline (range group) on a sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.index {number} The outline starting index.<br />
                 ** options.count {number} The number of rows or columns to group or ungroup in the outline.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var removeColumnOutline: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, index: number, count: number}, isUndo: boolean): any};
                /**
                 * Represents the command for ungrouping a row outline (range group) on a sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.index {number} The outline starting index.<br />
                 ** options.count {number} The number of rows or columns to group or ungroup in the outline.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var removeRowOutline: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, index: number, count: number}, isUndo: boolean): any};
                /**
                 * Represents the command used to rename a worksheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.name {string} The sheet's new name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var renameSheet: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, name: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to resize the column on a worksheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.columns {Array} The resize columns; each item is an object which has firstCol and lastCol.<br />
                 ** options.size {number} The size of the column that is being resized.<br />
                 ** options.rowHeader {boolean} Whether the column being resized is in the row header area.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var resizeColumn: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, columns: Object[], size: number, rowHeader: boolean}, isUndo: boolean): any};
                /**
                 * Represents the command for resizing floating objects.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.floatingObjects {Array} The names array of floating objects.<br />
                 ** options.offsetX The offset left.<br />
                 ** options.offsetY The offset top.<br />
                 ** options.offsetWidth The offset width.<br />
                 ** options.offsetHeight The offset height.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var resizeFloatingObjects: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, floatingObjects: string[], offsetX: number, offsetY: number, offsetWidth: number, offsetHeight: number}, isUndo: boolean): any};
                /**
                 * Represents the command used to resize the row on a worksheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.rows {Array} The resize rows; each item is an object which has firstRow and lastRow.<br />
                 ** options.size {number} The size of the row that is being resized.<br />
                 ** options.columnHeader {boolean} Whether the row being resized is in the column header area.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var resizeRow: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, rows: Object[], size: number, columnHeader: boolean}, isUndo: boolean): any};
                /**
                 * Represents the command used to extend the selection to the last row.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectionBottom: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to extend the selection down one row.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectionDown: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to extend the selection to the last column.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectionEnd: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to extend the selection to the first cell.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectionFirst: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to extend the selection to the first column.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectionHome: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to extend the selection to the last cell.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectionLast: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to extend the selection one column to the left.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectionLeft: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to extend the selection down to include one page of rows.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectionPageDown: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to extend the selection up to include one page of rows.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectionPageUp: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to extend the selection one column to the right.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectionRight: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to extend the selection to the first row.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectionTop: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to extend the selection up one row.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectionUp: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to select the next control.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectNextControl: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to select the previous control.
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute operation or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var selectPreviousControl: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to perform an undo of the most recent edit or action.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var undo: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string}, isUndo: boolean): any};
                /**
                 * Represents the command used to zoom the sheet.<br />
                 *
                 * The command has a boolean field <b>canUndo</b> that indicates whether the command supports undo and redo operations,
                 * and an <b>execute</b> method that performs an execute or undo operation.<br />
                 *
                 * The arguments of the execute method are as follows.<br />
                 * context {GC.Spread.Sheets.Workbook} The context of the operation.<br />
                 * options {Object} The options of the operation.<br />
                 ** options.sheetName {string} The sheet name.<br />
                 ** options.zoomFactor {number} The zoom factor.<br />
                 * isUndo {boolean} <c>true</c> if this is an undo operation; otherwise, <c>false</c>.
                 */
                var zoom: { canUndo: boolean, execute(context: GC.Spread.Sheets.Workbook, options: {sheetName: string, zoomFactor: number}, isUndo: boolean): any};
            }

            module Comments{
                /**
                 * Defines the comment state.
                 * @enum {number}
                 */
                export enum CommentState{
                    /**
                     * Specifies that the comment is in an active state.
                     */
                    active= 1,
                    /**
                     *Specifies that the comment is in an editing state.
                     */
                    edit= 2,
                    /**
                     * Specifies that the comment is in a normal state.
                     */
                    normal= 3
                }

                /**
                 * Defines when the comment is displayed.
                 * @enum {number}
                 */
                export enum DisplayMode{
                    /**
                     *  Specifies that the comment is always displayed.
                     */
                    alwaysShown= 1,
                    /**
                     *  Specifies that the comment is displayed only when the pointer hovers over the comment's owner cell.
                     */
                    hoverShown= 2
                }


                export class Comment{
                    /**
                     * Represents a comment.
                     * @class
                     * @param text The text of the comment.
                     */
                    constructor();
                    /**
                     * Gets or sets whether the comment automatically sizes based on its content.
                     * @param {boolean} value Whether the comment automatically sizes.
                     * @returns {boolean | GC.Spread.Sheets.Comments.Comment} If no value is set, returns whether the comment automatically sizes; otherwise, returns the comment.
                     */
                    autoSize(value?:  boolean): any;
                    /**
                     * Gets or sets the background color of the comment.
                     * @param {string} value The background color of the comment.
                     * @returns {string | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the background color of the comment; otherwise, returns the comment.
                     */
                    backColor(value?:  string): any;
                    /**
                     * Gets or sets the border color for the comment.
                     * @param {string} value The border color for the comment.
                     * @returns {string | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the border color for the comment; otherwise, returns the comment.
                     */
                    borderColor(value?:  string): any;
                    /**
                     * Gets or sets the border style for the comment.
                     * @param {string} value The border style for the comment.
                     * @returns {string | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the border style for the comment; otherwise, returns the comment.
                     */
                    borderStyle(value?:  string): any;
                    /**
                     * Gets or sets the border width for the comment.
                     * @param {number} value The border width for the comment.
                     * @returns {number | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the border width for the comment; otherwise, returns the comment.
                     */
                    borderWidth(value?:  number): any;
                    /**
                     * Gets or sets the state of the comment.
                     * @param {GC.Spread.Sheets.Comments.CommentState} value The state of the comment.
                     * @returns {GC.Spread.Sheets.Comments.CommentState | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the state of the comment; otherwise, returns the comment.
                     */
                    commentState(value?:  GC.Spread.Sheets.Comments.CommentState): any;
                    /**
                     * Gets or sets the display mode for the comment.
                     * @param {GC.Spread.Sheets.Comments.DisplayMode} value The display mode for the comment.
                     * @returns {GC.Spread.Sheets.Comments.DisplayMode | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the display mode for the comment; otherwise, returns the comment.
                     */
                    displayMode(value?:  GC.Spread.Sheets.Comments.DisplayMode): any;
                    /**
                     * Gets or sets whether the comment dynamically moves.
                     * @param {boolean} value Whether the comment dynamically moves.
                     * @returns {boolean | GC.Spread.Sheets.Comments.Comment} If no value is set, returns whether the comment dynamically moves; otherwise, returns the comment.
                     */
                    dynamicMove(value?:  boolean): any;
                    /**
                     * Gets or sets whether the comment is dynamically sized.
                     * @param {boolean} value Whether the comment is dynamically sized.
                     * @returns {boolean | GC.Spread.Sheets.Comments.Comment} If no value is set, returns whether the comment is dynamically sized; otherwise, returns the comment.
                     */
                    dynamicSize(value?:  boolean): any;
                    /**
                     * Gets or sets the font family for the comment.
                     * @param {string} value The font family for the comment.
                     * @returns {string | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the font family for the comment; otherwise, returns the comment.
                     */
                    fontFamily(value?:  string): any;
                    /**
                     * Gets or sets the font size for the comment. Valid value is numbers followed by "pt" (required), such as "12pt".
                     * @param {string} value The font size for the comment.
                     * @returns {string | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the font size for the comment; otherwise, returns the comment.
                     */
                    fontSize(value?:  string): any;
                    /**
                     * Gets or sets the font style of the comment.
                     * @param {string} value The font style of the comment.
                     * @returns {string | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the font style of the comment; otherwise, returns the comment.
                     */
                    fontStyle(value?:  string): any;
                    /**
                     * Gets or sets the font weight for the comment.
                     * @param {string} value The font weight for the comment.
                     * @returns {string | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the font weight for the comment; otherwise, returns the comment.
                     */
                    fontWeight(value?:  string): any;
                    /**
                     * Gets or sets the text color for the comment.
                     * @param {string} value The text color for the comment.
                     * @returns {string | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the text color for the comment; otherwise, returns the comment.
                     */
                    foreColor(value?:  string): any;
                    /**
                     * Gets or sets the height of the comment.
                     * @param {number} value The height of the comment.
                     * @returns {number | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the height of the comment; otherwise, returns the comment.
                     */
                    height(value?:  number): any;
                    /**
                     * Gets or sets the horizontal alignment of the comment.
                     * @param {GC.Spread.Sheets.HorizontalAlign} value The horizontal alignment of the comment.
                     * @returns {GC.Spread.Sheets.HorizontalAlign | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the horizontal alignment of the comment; otherwise, returns the comment.
                     */
                    horizontalAlign(value?:  GC.Spread.Sheets.HorizontalAlign): any;
                    /**
                     * Gets or sets the location of the comment.
                     * @param {GC.Spread.Sheets.Point} value The location of the comment.
                     * @returns {GC.Spread.Sheets.Point | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the location of the comment; otherwise, returns the comment.
                     */
                    location(value?:  GC.Spread.Sheets.Point): any;
                    /**
                     * Gets or sets the locked setting for the comment.
                     * @param {boolean} value The locked setting for the comment.
                     * @returns {boolean | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the locked setting for the comment; otherwise, returns the comment.
                     */
                    locked(value?:  boolean): any;
                    /**
                     * Gets or sets the locked text for the comment.
                     * @param {boolean} value The locked text for the comment.
                     * @returns {boolean | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the locked text for the comment; otherwise, returns the comment.
                     */
                    lockText(value?:  boolean): any;
                    /**
                     * Gets or sets the opacity of the comment.
                     * @param {number} value The opacity of the comment.
                     * @returns {number | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the opacity of the comment; otherwise, returns the comment.
                     */
                    opacity(value?:  number): any;
                    /**
                     * Gets or sets the padding for the comment.
                     * @param {GC.Spread.Sheets.Comments.Padding} value The padding for the comment.
                     * @returns {GC.Spread.Sheets.Comments.Padding | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the padding for the comment; otherwise, returns the comment.
                     */
                    padding(value?:  GC.Spread.Sheets.Comments.Padding): any;
                    /**
                     * Gets or sets whether the comment displays a shadow.
                     * @param {boolean} value Whether the comment displays a shadow.
                     * @returns {boolean | GC.Spread.Sheets.Comments.Comment} If no value is set, returns whether the comment displays a shadow; otherwise, returns the comment.
                     */
                    showShadow(value?:  boolean): any;
                    /**
                     * Gets or sets the text of the comment.
                     * @param {string} value The text of the comment.
                     * @returns {string | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the text of the comment; otherwise, returns the comment.
                     */
                    text(value?:  string): any;
                    /**
                     * Gets or sets the text decoration for the comment.
                     * @param {GC.Spread.Sheets.TextDecorationType} value The text decoration for the comment.
                     * @returns {GC.Spread.Sheets.TextDecorationType | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the text decoration for the comment; otherwise, returns the comment.
                     */
                    textDecoration(value?:  GC.Spread.Sheets.TextDecorationType): any;
                    /**
                     * Gets or sets the width of the comment.
                     * @param {number} value The width of the comment.
                     * @returns {number | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the width of the comment; otherwise, returns the comment.
                     */
                    width(value?:  number): any;
                    /**
                     * Gets or sets the z-index of the comment.
                     * @param {number} value The z-index of the comment.
                     * @returns {number | GC.Spread.Sheets.Comments.Comment} If no value is set, returns the z-index of the comment; otherwise, returns the comment.
                     */
                    zIndex(value?:  number): any;
                }

                export class CommentManager{
                    /**
                     * Represents a comment manager that can manage all comments in a sheet.
                     * @class
                     * @param {GC.Spread.Sheets.Worksheet} sheet The worksheet.
                     */
                    constructor(sheet:  GC.Spread.Sheets.Worksheet);
                    /**
                     * Adds a comment to the cell for the indicated row and column.
                     * @param {number} row The row index of the cell.
                     * @param {number} col The column index of the cell.
                     * @param {string} text The text of the comment.
                     * @returns {GC.Spread.Sheets.Comments.Comment} The comment that has been added to the cell.
                     */
                    add(row:  number,  col:  number,  text:  string): GC.Spread.Sheets.Comments.Comment;
                    /**
                     * Gets all comments in the sheet.
                     * @returns {Array.<GC.Spread.Sheets.Comments.Comment>}
                     */
                    all(): GC.Spread.Sheets.Comments.Comment[];
                    /**
                     * Clears all of the comments in the indicated range on the sheet. When the range is not specified, it clears all the comments in the sheet.
                     * @param {GC.Spread.Sheets.Range} range The range that you want clear all comments from.
                     */
                    clear(range:  GC.Spread.Sheets.Range): void;
                    /**
                     * Gets the comment in the cell with the indicated row and column.
                     * @param {number} row The row index of the cell.
                     * @param {number} col The column index of the cell.
                     * @returns {GC.Spread.Sheets.Comments.Comment} The comment in the indicated cell.
                     */
                    get(row:  number,  col:  number): GC.Spread.Sheets.Comments.Comment;
                    /**
                     *Removes the comment from the cell for the indicated row and column.
                     * @param {number} row The row index of the cell.
                     * @param {number} col The column index of the cell.
                     */
                    remove(row:  number,  col:  number): void;
                }

                export class Padding{
                    /**
                     * Represents the padding information.
                     * @class
                     * @param {number} top The top padding.
                     * @param {number} right The right padding.
                     * @param {number} bottom The bottom padding.
                     * @param {number} left The left padding.
                     */
                    constructor(top?:  number,  right?:  number,  bottom?:  number,  left?:  number);
                }
            }

            module ConditionalFormatting{
                /**
                 * Specifies the average condition type.
                 * @enum {number}
                 */
                export enum AverageConditionType{
                    /** Specifies the above average condition.
                     * @type {number}
                     */
                    above= 0,
                    /** Specifies the below average condition.
                     * @type {number}
                     */
                    below= 1,
                    /** Specifies the above average or equal average condition.
                     * @type {number}
                     */
                    equalOrAbove= 2,
                    /** Specifies the below average or equal average condition.
                     * @type {number}
                     */
                    equalOrBelow= 3,
                    /** Specifies the above standard deviation condition.
                     * @type {number}
                     */
                    above1StdDev= 4,
                    /** Specifies the below standard deviation condition.
                     * @type {number}
                     */
                    below1StdDev= 5,
                    /** Specifies above the two standard deviation condition.
                     * @type {number}
                     */
                    above2StdDev= 6,
                    /** Specifies below the two standard deviation condition.
                     * @type {number}
                     */
                    below2StdDev= 7,
                    /** Specifies above the three standard deviation condition.
                     * @type {number}
                     */
                    above3StdDev= 8,
                    /** Specifies below the three standard deviation condition.
                     * @type {number}
                     */
                    below3StdDev= 9
                }

                /**
                 * Specifies the data bar direction.
                 * @enum {number}
                 */
                export enum BarDirection{
                    /** Specifies whether to display the data bar from left to right.
                     * @type {number}
                     */
                    leftToRight= 0,
                    /** Specifies whether to display the data bar from right to left.
                     * @type {number}
                     */
                    rightToLeft= 1
                }

                /**
                 * Specifies the color compare type.
                 * @enum {number}
                 */
                export enum ColorCompareType{
                    /** Indicates whether the cell background color is equal to a specified color.
                     * @type {number}
                     */
                    backgroundColor= 0,
                    /** Indicates whether the cell foreground color is equal to a specified color.
                     * @type {number}
                     */
                    foregroundColor= 1
                }

                /**
                 * Specifies the comparison operator.
                 * @enum {number}
                 */
                export enum ComparisonOperators{
                    /** Determines whether a cell value is equal to the parameter value.
                     * @type {number}
                     */
                    equalsTo= 0,
                    /** Determines whether a cell value is not equal to the parameter value.
                     * @type {number}
                     */
                    notEqualsTo= 1,
                    /** Determines whether a cell value is greater than the parameter value.
                     * @type {number}
                     */
                    greaterThan= 2,
                    /** Determines whether a cell value is greater than or equal to the parameter value.
                     * @type {number}
                     */
                    greaterThanOrEqualsTo= 3,
                    /** Determines whether a cell value is less than the parameter value.
                     * @type {number}
                     */
                    lessThan= 4,
                    /** Determines whether a cell value is less than or equal to the parameter value.
                     * @type {number}
                     */
                    lessThanOrEqualsTo= 5,
                    /** Determines whether a cell value is between the two parameter values.
                     * @type {number}
                     */
                    between= 6,
                    /** Determines whether a cell value is not between the two parameter values.
                     * @type {number}
                     */
                    notBetween= 7
                }

                /**
                 * Specifies the condition type.
                 * @enum {number}
                 */
                export enum ConditionType{
                    /** Specifies the relation condition.
                     * @type {number}
                     */
                    relationCondition= 0,
                    /** Specifies the number condition.
                     * @type {number}
                     */
                    numberCondition= 1,
                    /** Specifies the text condition.
                     * @type {number}
                     */
                    textCondition= 2,
                    /** Specifies the color condition.
                     * @type {number}
                     */
                    colorCondition= 3,
                    /** Specifies the formula condition.
                     * @type {number}
                     */
                    formulaCondition= 4,
                    /** Specifies the date condition.
                     * @type {number}
                     */
                    dateCondition= 5,
                    /** Specifies the dateex condition.
                     * @type {number}
                     */
                    dateExCondition= 6,
                    /** Specifies the text length condition.
                     * @type {number}
                     */
                    textLengthCondition= 7,
                    /** Specifies the top 10 condition.
                     * @type {number}
                     */
                    top10Condition= 8,
                    /** Specifies the unique condition.
                     * @type {number}
                     */
                    uniqueCondition= 9,
                    /** Specifies the average condition.
                     * @type {number}
                     */
                    averageCondition= 10,
                    /** Specifies the cell value condition.
                     * @type {number}
                     */
                    cellValueCondition= 11,
                    /** Specifies the area condition.
                     * @type {number}
                     */
                    areaCondition= 12
                }

                /**
                 * Specifies the custom value type.
                 * @enum {number}
                 */
                export enum CustomValueType{
                    /** Indicates whether the cell value is empty or null.
                     * @type {number}
                     */
                    empty= 0,
                    /** Indicates whether the cell value is not empty or null.
                     * @type {number}
                     */
                    nonEmpty= 1,
                    /** Indicates whether the cell value contains a calculation error.
                     * @type {number}
                     */
                    error= 2,
                    /** Indicates whether the cell value does not contain a calculation error.
                     * @type {number}
                     */
                    nonError= 3,
                    /** Indicates whether the cell value is a formula.
                     * @type {number}
                     */
                    formula= 4
                }

                /**
                 * Specifies the position of the data bar's axis.
                 * @enum {number}
                 */
                export enum DataBarAxisPosition{
                    /** Specifies whether to display at a variable position based on negative values.
                     * @type {number}
                     */
                    automatic= 0,
                    /** Specifies whether to display at the cell midpoint.
                     * @type {number}
                     */
                    cellMidPoint= 1,
                    /** Specifies whether to display value bars in the same direction as positive values.
                     * @type {number}
                     */
                    none= 2
                }

                /**
                 * Specifies the date compare type.
                 * @enum {number}
                 */
                export enum DateCompareType{
                    /** Indicates whether the date time is equal to a certain time.
                     * @type {number}
                     */
                    equalsTo= 0,
                    /** Indicates whether the date time is not equal to a certain time.
                     * @type {number}
                     */
                    notEqualsTo= 1,
                    /** Indicates whether the date time is before a certain time.
                     * @type {number}
                     */
                    before= 2,
                    /** Indicates whether the date time is before or equal to a certain time.
                     * @type {number}
                     */
                    beforeEqualsTo= 3,
                    /** Indicates whether the date time is after a certain time.
                     * @type {number}
                     */
                    after= 4,
                    /** Indicates whether the date time is after or equal to a certain time.
                     * @type {number}
                     */
                    afterEqualsTo= 5
                }

                /**
                 * Specifies the date occurring type.
                 * @enum {number}
                 */
                export enum DateOccurringType{
                    /** Specifies today.
                     * @type {number}
                     */
                    today= 0,
                    /** Specifies yesterday.
                     * @type {number}
                     */
                    yesterday= 1,
                    /** Specifies tomorrow.
                     * @type {number}
                     */
                    tomorrow= 2,
                    /** Specifies the last seven days.
                     * @type {number}
                     */
                    last7Days= 3,
                    /** Specifies this month.
                     * @type {number}
                     */
                    thisMonth= 4,
                    /** Specifies last month.
                     * @type {number}
                     */
                    lastMonth= 5,
                    /** Specifies next month.
                     * @type {number}
                     */
                    nextMonth= 6,
                    /** Specifies this week.
                     * @type {number}
                     */
                    thisWeek= 7,
                    /** Specifies last week.
                     * @type {number}
                     */
                    lastWeek= 8,
                    /** Specifies next week.
                     * @type {number}
                     */
                    nextWeek= 9
                }

                /**
                 * Specifies the general operator.
                 * @enum {number}
                 */
                export enum GeneralComparisonOperators{
                    /** Indicates whether the number is equal to a specified number.
                     * @type {number}
                     */
                    equalsTo= 0,
                    /** Indicates whether the number is not equal to a specified number.
                     * @type {number}
                     */
                    notEqualsTo= 1,
                    /** Indicates whether the number is greater than a specified number.
                     * @type {number}
                     */
                    greaterThan= 2,
                    /** Indicates whether the number is greater than or equal to a specified number.
                     * @type {number}
                     */
                    greaterThanOrEqualsTo= 3,
                    /** Indicates whether the number is less than a specified number.
                     * @type {number}
                     */
                    lessThan= 4,
                    /** Indicates whether the number is less than or equal to a specified number.
                     * @type {number}
                     */
                    lessThanOrEqualsTo= 5
                }

                /**
                 * Specifies the icon set.
                 * @enum {number}
                 */
                export enum IconSetType{
                    /** Specifies three colored arrows.
                     * @type {number}
                     */
                    threeArrowsColored= 0,
                    /** Specifies three gray arrows.
                     * @type {number}
                     */
                    threeArrowsGray= 1,
                    /** Specifies three trangles.
                     * @type {number}
                     */
                    threeTriangles= 2,
                    /** Specifies three stars.
                     * @type {number}
                     */
                    threeStars= 3,
                    /** Specifies three flags.
                     * @type {number}
                     */
                    threeFlags= 4,
                    /** Specifies three traffic lights (unrimmed).
                     * @type {number}
                     */
                    threeTrafficLightsUnrimmed= 5,
                    /** Specifies three traffic lights (rimmed).
                     * @type {number}
                     */
                    threeTrafficLightsRimmed= 6,
                    /** Specifies three signs.
                     * @type {number}
                     */
                    threeSigns= 7,
                    /** Specifies three symbols (circled).
                     * @type {number}
                     */
                    threeSymbolsCircled= 8,
                    /** Specifies three symbols (uncircled).
                     * @type {number}
                     */
                    threeSymbolsUncircled= 9,
                    /** Specifies four colored arrows.
                     * @type {number}
                     */
                    fourArrowsColored= 10,
                    /** Specifies four gray arrows.
                     * @type {number}
                     */
                    fourArrowsGray= 11,
                    /** Specifies four red to black icons.
                     * @type {number}
                     */
                    fourRedToBlack= 12,
                    /** Specifies four ratings.
                     * @type {number}
                     */
                    fourRatings= 13,
                    /** Specifies four traffic lights.
                     * @type {number}
                     */
                    fourTrafficLights= 14,
                    /** Specifies five colored arrows.
                     * @type {number}
                     */
                    fiveArrowsColored= 15,
                    /** Specifies five gray arrows.
                     * @type {number}
                     */
                    fiveArrowsGray= 16,
                    /** Specifies five ratings.
                     * @type {number}
                     */
                    fiveRatings= 17,
                    /** Specifies five quarters.
                     * @type {number}
                     */
                    fiveQuarters= 18,
                    /** Specifies five boxes.
                     * @type {number}
                     */
                    fiveBoxes= 19
                }

                /**
                 * Specifies the icon value type.
                 * @enum {number}
                 */
                export enum IconValueType{
                    /** Indicates whether to return a specified number directly.
                     * @type {number}
                     */
                    number= 1,
                    /** Indicates whether to return the percentage of a cell value in a specified cell range.
                     * @type {number}
                     */
                    percent= 4,
                    /** Indicates whether to return the result of a formula calculation.
                     * @type {number}
                     */
                    formula= 7,
                    /** Indicates whether to return the percentile of a cell value in a specified cell range.
                     * @type {number}
                     */
                    percentile= 5
                }

                /**
                 * Specifies the relation operator.
                 * @enum {number}
                 */
                export enum LogicalOperators{
                    /** Specifies the Or relation.
                     * @type {number}
                     */
                    or= 0,
                    /** Specifies the And relation.
                     * @type {number}
                     */
                    and= 1
                }

                /**
                 * Specifies the quarter type.
                 * @enum {number}
                 */
                export enum QuarterType{
                    /** Indicates the first quarter of a year.
                     * @type {number}
                     */
                    quarter1= 0,
                    /** Indicates the second quarter of a year.
                     * @type {number}
                     */
                    quarter2= 1,
                    /** Indicates the third quarter of a year.
                     * @type {number}
                     */
                    quarter3= 2,
                    /** Indicates the fourth quarter of a year.
                     * @type {number}
                     */
                    quarter4= 3
                }

                /**
                 * Specifies the rule type.
                 * @enum {number}
                 */
                export enum RuleType{
                    /** Specifies the base rule of the condition.
                     * @type {number}
                     */
                    conditionRuleBase= 0,
                    /** Specifies the cell value rule.
                     * @type {number}
                     */
                    cellValueRule= 1,
                    /** Specifies the specific text rule.
                     * @type {number}
                     */
                    specificTextRule= 2,
                    /** Specifies the formula rule.
                     * @type {number}
                     */
                    formulaRule= 3,
                    /** Specifies the date occurring rule.
                     * @type {number}
                     */
                    dateOccurringRule= 4,
                    /** Specifies the top 10 rule.
                     * @type {number}
                     */
                    top10Rule= 5,
                    /** Specifies the unique rule.
                     * @type {number}
                     */
                    uniqueRule= 6,
                    /** Specifies the duplicate rule.
                     * @type {number}
                     */
                    duplicateRule= 7,
                    /** Specifies the average rule.
                     * @type {number}
                     */
                    averageRule= 8,
                    //ScaleRule: 9,
                     /** Specifies the two scale rule.
                     * @type {number}
                     */
                    twoScaleRule= 10,
                    /** Specifies the three scale rule.
                     * @type {number}
                     */
                    threeScaleRule= 11,
                    /** Specifies the data bar rule.
                     * @type {number}
                     */
                    dataBarRule= 12,
                    /** Specifies the icon set rule.
                     * @type {number}
                     */
                    iconSetRule= 13
                }

                /**
                 * Specifies the scale value type.
                 * @enum {number}
                 */
                export enum ScaleValueType{
                    /** Indicates whether to return a specified number directly.
                     * @type {number}
                     */
                    number= 0,
                    /** Indicates whether to return the lowest value in a specified cell range.
                     * @type {number}
                     */
                    lowestValue= 1,
                    /** Indicates whether to return the highest value in a specified cell range.
                     * @type {number}
                     */
                    highestValue= 2,
                    /** Indicates whether to return the percentage of a cell value in a specified cell range.
                     * @type {number}
                     */
                    percent= 3,
                    /** Indicates whether to return the percentile of a cell value in a specified cell range.
                     * @type {number}
                     */
                    percentile= 4,
                    /** Indicates whether to return the automatic minimum value in a specified range.
                     * @type {number}
                     */
                    automin= 5,
                    /** Indicates whether to return the result of a formula calculation.
                     * @type {number}
                     */
                    formula= 6,
                    /** Indicates whether to return the automatic maximum value in a specified range.
                     * @type {number}
                     */
                    automax= 7
                }

                /**
                 * Specifies the text compare type.
                 * @enum {number}
                 */
                export enum TextCompareType{
                    /** Indicates whether the string is equal to a specified string.
                     * @type {number}
                     */
                    equalsTo= 0,
                    /** Indicates whether the string is not equal to a specified string.
                     * @type {number}
                     */
                    notEqualsTo= 1,
                    /** Indicates whether the string starts with a specified string.
                     * @type {number}
                     */
                    beginsWith= 2,
                    /** Indicates whether the string does not start with a specified string.
                     * @type {number}
                     */
                    doesNotBeginWith= 3,
                    /** Indicates whether the string ends with a specified string.
                     * @type {number}
                     */
                    endsWith= 4,
                    /** Indicates whether the string does not end with a specified string.
                     * @type {number}
                     */
                    doesNotEndWith= 5,
                    /** Indicates whether the string contains a specified string.
                     * @type {number}
                     */
                    contains= 6,
                    /** Indicates whether the string does not contain a specified string.
                     * @type {number}
                     */
                    doesNotContain= 7
                }

                /**
                 * Specifies the text comparison operator.
                 * @enum {number}
                 */
                export enum TextComparisonOperators{
                    /** Determines whether a cell value contains the parameter value.
                     * @type {number}
                     */
                    contains= 0,
                    /** Determines whether a cell value does not contain the parameter value.
                     * @type {number}
                     */
                    doesNotContain= 1,
                    /** Determines whether a cell value begins with the parameter value.
                     * @type {number}
                     */
                    beginsWith= 2,
                    /** Determines whether a cell value ends with the parameter value.
                     * @type {number}
                     */
                    endsWith= 3
                }

                /**
                 * Specifies the top 10 condition type.
                 * @enum {number}
                 */
                export enum Top10ConditionType{
                    /** Specifies the top condition.
                     * @type {number}
                     */
                    top= 0,
                    /** Specifies the bottom condition.
                     * @type {number}
                     */
                    bottom= 1
                }


                export class Condition{
                    /**
                     * Represents a conditional item using the parameter object.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ConditionType} conditionType
                     * @param {Object} args
                     * @constructor
                     */
                    constructor(conditionType:  ConditionType,  args:  Object);
                    /**
                     * Gets or sets the rule compare type.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.LogicalOperators | GC.Spread.Sheets.ConditionalFormatting.GeneralComparisonOperators | GC.Spread.Sheets.ConditionalFormatting.TextCompareType | GC.Spread.Sheets.ConditionalFormatting.ColorCompareType | GC.Spread.Sheets.ConditionalFormatting.DateCompareType} value The rule compare type.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.LogicalOperators | GC.Spread.Sheets.ConditionalFormatting.GeneralComparisonOperators | GC.Spread.Sheets.ConditionalFormatting.TextCompareType | GC.Spread.Sheets.ConditionalFormatting.ColorCompareType | GC.Spread.Sheets.ConditionalFormatting.DateCompareType | GC.Spread.Sheets.ConditionalFormatting.Condition} If no value is set, returns the rule compare type; otherwise, returns the condition.
                     */
                    compareType(value?:  any): any;
                    /**
                     * Evaluates the condition using the specified evaluator.
                     * @param {object} evaluator The evaluator that can evaluate an expression or a function.
                     * @param {number} baseRow The base row index for evaluation.
                     * @param {number} baseColumn The base column index for evaluation.
                     * @param {object} actualObj The actual value of object1 for evaluation.
                     * @returns {boolean} <c>true</c> if the result is successful; otherwise, <c>false</c>.
                     */
                    evaluate(evaluator:  Object,  baseRow:  number,  baseColumn:  number,  actualObj:  Object): boolean;
                    /**
                     * Gets or sets the expected value.
                     * @param {object} value The expected value.
                     * @returns {object | GC.Spread.Sheets.ConditionalFormatting.Condition} If no value is set, returns the expected value; otherwise, returns the condition.
                     */
                    expected(value?:  any): any;
                    /**
                     * Gets or sets the expected formula.
                     * @param {string} value The expected formula.
                     * @returns {string | GC.Spread.Sheets.ConditionalFormatting.Condition} If no value is set, returns the expected formula; otherwise, returns the condition.
                     */
                    formula(value?:  string): any;
                    /**
                     * Creates a date extend condition object from the specified day.
                     * @static
                     * @param {number} day The day.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.Condition} A date extend condition object.
                     */
                    static fromDay(day:  number): GC.Spread.Sheets.ConditionalFormatting.Condition;
                    /**
                     * Creates the area condition from formula data.
                     * @static
                     * @param {string} formula The formula that specifies a range that contains data items.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.Condition} The area condition.
                     */
                    static fromFormula(formula:  string): GC.Spread.Sheets.ConditionalFormatting.Condition;
                    /**
                     * Creates a date extend condition object from the specified month.
                     * @static
                     * @param {number} month The month. The first month is 0.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.Condition} A date extend condition object.
                     */
                    static fromMonth(month:  number): GC.Spread.Sheets.ConditionalFormatting.Condition;
                    /**
                     * Creates a date extend condition object from the specified quarter.
                     * @static
                     * @param {GC.Spread.Sheets.ConditionalFormatting.QuarterType} quarter The quarter.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.Condition} A date extend condition object.
                     */
                    static fromQuarter(quarter:  QuarterType): GC.Spread.Sheets.ConditionalFormatting.Condition;
                    /**
                     * Creates the area condition from source data.
                     * @static
                     * @param {string} expected The expected source that separates each data item with a comma (",").
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.Condition} The area condition.
                     */
                    static fromSource(expected:  string): GC.Spread.Sheets.ConditionalFormatting.Condition;
                    /**
                     * Creates a date extend condition object from the specified week.
                     * @static
                     * @param {number} week The week.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.Condition} A date extend condition object.
                     */
                    static fromWeek(week:  number): GC.Spread.Sheets.ConditionalFormatting.Condition;
                    /**
                     * Creates a date extend condition object from the specified year.
                     * @static
                     * @param {number} year The year.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.Condition} A date extend condition object.
                     */
                    static fromYear(year:  number): GC.Spread.Sheets.ConditionalFormatting.Condition;
                    /**
                     * Gets the expected value.
                     * @constructor
                     * @param {object} evaluator The evaluator that can evaluate an expression or a function.
                     * @param {number} baseRow The base row index for evaluation.
                     * @param {number} baseColumn The base column index for evaluation.
                     * @returns {object} The expected value.
                     */
                    getExpected(evaluator:  Object,  baseRow:  number,  baseColumn:  number): Object;
                    /**
                     * Returns the list of valid data items.
                     * @param {object} evaluator The evaluator that can evaluate an expression or a function.
                     * @param {number} baseRow The base row index for evaluation.
                     * @param {number} baseColumn The base column index for evaluation.
                     * @returns {Array} The list of valid data items.
                     */
                    getValidList(evaluator:  Object,  baseRow:  number,  baseColumn:  number): any[];
                    /**
                     * Gets or sets whether to ignore the blank cell.
                     * @param {boolean} value Whether to ignore the blank cell.
                     * @returns {boolean | GC.Spread.Sheets.ConditionalFormatting.Condition} If no value is set, returns whether to ignore the blank cell; otherwise, returns the condition.
                     */
                    ignoreBlank(value?:  boolean): any;
                    /**
                     * Gets or sets whether to ignore case when performing the comparison.
                     * @param {boolean} value Whether to ignore case when performing the comparison.
                     * @returns {boolean | GC.Spread.Sheets.ConditionalFormatting.Condition} If no value is set, returns whether to ignore case when performing the comparison; otherwise, returns the condition.
                     */
                    ignoreCase(value?:  boolean): any;
                    /**
                     * Gets or sets the first condition.
                     * @param {object} value The first condition.
                     * @returns {object | GC.Spread.Sheets.ConditionalFormatting.Condition} If no value is set, returns the first condition; otherwise, returns the relation condition.
                     */
                    item1(value?:  Object): any;
                    /**
                     * Gets or sets the second condition.
                     * @param {object} value The second condition.
                     * @returns {object | GC.Spread.Sheets.ConditionalFormatting.Condition} If no value is set, returns the second condition; otherwise, returns the relation condition.
                     */
                    item2(value?:  Object): any;
                    /**
                     * Gets or sets the condition ranges.
                     * @param {Array.<GC.Spread.Sheets.Range>} value The condition ranges.
                     * @returns {Array.<GC.Spread.Sheets.Range> | GC.Spread.Sheets.ConditionalFormatting.Condition} If no value is set, returns the condition ranges; otherwise, returns the condition.
                     */
                    ranges(value?:  GC.Spread.Sheets.Range[]): any;
                    /**
                     * Resets this instance.
                     */
                    reset(): void;
                    /**
                     * Gets or sets whether to treat the null value in a cell as zero.
                     * @param {boolean} value Whether to treat the null value in a cell as zero.
                     * @returns {boolean | GC.Spread.Sheets.ConditionalFormatting.Condition} If no value is set, returns whether to treat the null value in a cell as zero; otherwise, returns the condition.
                     */
                    treatNullValueAsZero(value?:  boolean): any;
                    /**
                     * Gets or sets whether to compare strings using wildcards.
                     * @param {boolean} value Whether to compare strings using wildcards.
                     * @returns {boolean | GC.Spread.Sheets.ConditionalFormatting.Condition} If no value is set, returns whether to compare strings using wildcards; otherwise, returns the condition.
                     */
                    useWildCards(value?:  boolean): any;
                }

                export class ConditionalFormats{
                    /**
                     * Represents a format condition class.
                     * @class
                     * @param {object} worksheet The sheet.
                     */
                    constructor(worksheet:  GC.Spread.Sheets.Worksheet);
                    /**
                     * Adds the two scale rule to the rule collection.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} minType The minimum scale type.
                     * @param {object} minValue The minimum scale value.
                     * @param {string} minColor The minimum scale color.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} maxType The maximum scale type.
                     * @param {object} maxValue The maximum scale value.
                     * @param {string} maxColor The maximum scale color.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell ranges where the rule is applied whose item type is GC.Spread.Sheets.Range.
                     * @returns {object} The two scale rule added to the rule collection.
                     */
                    add2ScaleRule(minType:  ScaleValueType,  minValue:  Object,  minColor:  string,  maxType:  ScaleValueType,  maxValue:  Object,  maxColor:  string,  ranges:  GC.Spread.Sheets.Range[]): Object;
                    /**
                     * Adds the three scale rule to the rule collection.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} minType The minimum scale type.
                     * @param {object} minValue The minimum scale value.
                     * @param {string} minColor The minimum scale color.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} midType The midpoint scale type.
                     * @param {object} midValue The midpoint scale value.
                     * @param {string} midColor The midpoint scale color.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} maxType The maximum scale type.
                     * @param {object} maxValue The maximum scale value.
                     * @param {string} maxColor The maximum scale color.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell ranges where the rule is applied whose item type is GC.Spread.Sheets.Range.
                     * @returns {object} The three scale rule added to the rule collection.
                     */
                    add3ScaleRule(minType:  ScaleValueType,  minValue:  Object,  minColor:  string,  midType:  ScaleValueType,  midValue:  Object,  midColor:  string,  maxType:  ScaleValueType,  maxValue:  Object,  maxColor:  string,  ranges:  GC.Spread.Sheets.Range[]): Object;
                    /**
                     * Adds an average rule to the rule collection.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.AverageConditionType} type The average condition type.
                     * @param {GC.Spread.Sheets.Style} style The style that is applied to the cell when the condition is met.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell ranges where the rule is applied whose item type is GC.Spread.Sheets.Range.
                     * @returns {object} The average rule added to the rule collection.
                     */
                    addAverageRule(type:  AverageConditionType,  style:  GC.Spread.Sheets.Style,  ranges:  GC.Spread.Sheets.Range[]): Object;
                    /**
                     * Adds the cell value rule to the rule collection.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators} comparisionOperator The comparison operator.
                     * @param {object} value1 The first value.
                     * @param {object} value2 The second value.
                     * @param {GC.Spread.Sheets.Style} style The style that is applied to the cell when the condition is met.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell ranges where the rule is applied whose item type is GC.Spread.Sheets.Range.
                     * @returns {object} The cell value rule added to the rule collection.
                     */
                    addCellValueRule(comparisionOperator:  ComparisonOperators,  value1:  Object,  value2:  Object,  style:  GC.Spread.Sheets.Style,  ranges:  GC.Spread.Sheets.Range[]): Object;
                    /**
                     * Adds a data bar rule to the rule collection.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} minType The minimum scale type.
                     * @param {object} minValue The minimum scale value.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} maxType The maximum scale type.
                     * @param {object} maxValue The maximum scale value.
                     * @param {string} color The color data bar to show on the view.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell ranges where the rule is applied whose item type is GC.Spread.Sheets.Range.
                     * @returns {object} The data bar rule added to the rule collection.
                     */
                    addDataBarRule(minType:  ScaleValueType,  minValue:  Object,  maxType:  ScaleValueType,  maxValue:  Object,  color:  string,  ranges:  GC.Spread.Sheets.Range[]): Object;
                    /**
                     * Adds the date occurring rule to the rule collection.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.DateOccurringType} type The data occurring type.
                     * @param {GC.Spread.Sheets.Style} style The style that is applied to the cell when the condition is met.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell ranges where the rule is applied whose item type is GC.Spread.Sheets.Range.
                     * @returns {object} The date occurring rule added to the rule collection.
                     */
                    addDateOccurringRule(type:  DateOccurringType,  style:  GC.Spread.Sheets.Style,  ranges:  GC.Spread.Sheets.Range[]): Object;
                    /**
                     * Adds a duplicate rule to the rule collection.
                     * @param {GC.Spread.Sheets.Style} style The style that is applied to the cell when the condition is met.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell ranges where the rule is applied whose item type is GC.Spread.Sheets.Range.
                     * @returns {object} The duplicate rule added to the rule collection.
                     */
                    addDuplicateRule(style:  GC.Spread.Sheets.Style,  ranges:  GC.Spread.Sheets.Range[]): Object;
                    /**
                     * Adds the formula rule to the rule collection.
                     * @param {string} formula The condition formula.
                     * @param {GC.Spread.Sheets.Style} style The style that is applied to the cell when the condition is met.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell ranges where the rule is applied whose item type is GC.Spread.Sheets.Range.
                     * @returns {object} The formula rule added to the rule collection.
                     */
                    addFormulaRule(formula:  string,  style:  GC.Spread.Sheets.Style,  ranges:  GC.Spread.Sheets.Range[]): Object;
                    /**
                     * Adds an icon set rule to the rule collection.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.IconSetType} iconSetTye The type of icon set.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell ranges where the rule is applied whose item type is GC.Spread.Sheets.Range.
                     * @returns {object} The icon set rule added to the rule collection.
                     */
                    addIconSetRule(iconSetTye:  IconSetType,  ranges:  GC.Spread.Sheets.Range[]): Object;
                    /**
                     * Adds the rule.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase} rule The rule to add.
                     */
                    addRule(rule:  ConditionRuleBase): void;
                    /**
                     * Adds the text rule to the rule collection.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.TextComparisonOperators} comparisonOperator The comparison operator.
                     * @param {string} text The text for comparison.
                     * @param {GC.Spread.Sheets.Style} style The style that is applied to the cell when the condition is met.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell ranges where the rule is applied to items whose item type is GC.Spread.Sheets.Range.
                     * @returns {object} The text rule added to the rule collection.
                     */
                    addSpecificTextRule(comparisionOperator:  TextComparisonOperators,  text:  string,  style:  GC.Spread.Sheets.Style,  ranges:  GC.Spread.Sheets.Range[]): Object;
                    /**
                     * Adds the top 10 rule or bottom 10 rule to the collection based on the Top10CondtionType object.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.Top10ConditionType} type The top 10 condition.
                     * @param {number} rank The number of top or bottom items to apply the style to.
                     * @param {GC.Spread.Sheets.Style} style The style that is applied to the cell when the condition is met.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell ranges where the rule is applied whose item type is GC.Spread.Sheets.Range.
                     * @returns {object} The top 10 rule added to the rule collection.
                     */
                    addTop10Rule(type:  Top10ConditionType,  rank:  number,  style:  GC.Spread.Sheets.Style,  ranges:  GC.Spread.Sheets.Range[]): Object;
                    /**
                     * Adds a unique rule to the rule collection.
                     * @param {GC.Spread.Sheets.Style} style The style that is applied to the cell when the condition is met.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell ranges where the rule is applied whose item type is GC.Spread.Sheets.Range.
                     * @returns {object} The unique rule added to the rule collection.
                     */
                    addUniqueRule(style:  GC.Spread.Sheets.Style,  ranges:  GC.Spread.Sheets.Range[]): Object;
                    /**
                     * Removes all rules.
                     */
                    clearRule(): void;
                    /**
                     * Determines whether the specified cell contains a specified rule.
                     * @param {object} rule The rule for which to check.
                     * @param {number} row The row index.
                     * @param {number} column The column index.
                     * @returns {boolean} <c>true</c> if the specified cell contains a specified rule; otherwise, <c>false</c>.
                     */
                    containsRule(rule:  Object,  row:  number,  column:  number): boolean;
                    /**
                     * Gets the number of rule objects in the collection.
                     * @returns {number} The number of rule objects in the collection.
                     */
                    count(): number;
                    /**
                     * Gets the rule using the index.
                     * @param {number} index The index from which to get the rule.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase} The rule from the index.
                     */
                    getRule(index:  number): ConditionRuleBase;
                    /**
                     * Gets the conditional rules from the cell at the specified row and column.
                     * @param {number} row The row index.
                     * @param {number} column The column index.
                     * @returns {Array.<GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase>} The conditional rules.
                     */
                    getRules(row:  number,  column:  number): ConditionRuleBase[];
                    /**
                     * Removes a rule object from the ConditionalFormats object.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase} rule The rule object to remove from the ConditionalFormats object.
                     */
                    removeRule(rule:  ConditionRuleBase): void;
                    /**
                     * Removes the rules from a specified cell range.
                     * @param {number} row The row index of the first cell in the range.
                     * @param {number} column The column index of the first cell in the range.
                     * @param {number} rowCount The number of rows in the range.
                     * @param {number} columnCount The number of columns in the range.
                     */
                    removeRuleByRange(row:  number,  column:  number,  rowCount:  number,  columnCount:  number): void;
                }

                export class ConditionRuleBase{
                    /**
                     * Represents a formatting base rule class as the specified style.
                     * @param ruleType
                     * @param {GC.Spread.Sheets.Style} style The style for the rule.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The range array.
                     * @class
                     */
                    constructor(ruleType:  RuleType,  style:  GC.Spread.Sheets.Style,  ranges:  GC.Spread.Sheets.Range[]);
                    /**
                     * Gets or sets the base condition of the rule.
                     * @param {object} value The base condition of the rule.
                     * @returns {object | GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase} If no value is set, returns the base condition of the rule; otherwise, returns the condition rule.
                     */
                    condition(value?:  Condition): any;
                    /**
                     * Determines whether the range of cells contains the cell at the specified row and column.
                     * @param {number} row The row index.
                     * @param {number} column The column index.
                     * @returns {boolean} <c>true</c> if the range of cells contains the cell at the specified row and column; otherwise, <c>false</c>.
                     */
                    contains(row:  number,  column:  number): boolean;
                    /**
                     * Creates condition for the rule.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.Condition} The condition.
                     */
                    createCondition(): Condition;
                    /**
                     * Returns the cell style of the rule if the cell satisfies the condition.
                     * @param {object} evaluator The object that can evaluate a condition.
                     * @param {number} baseRow The row index.
                     * @param {number} baseColumn The column index.
                     * @param {object} actual The actual value.
                     * @returns {object} The cell style of the rule.
                     */
                    evaluate(evaluator:  Object,  baseRow:  number,  baseColumn:  number,  actual:  Object): Object;
                    /**
                     * Gets the style of the base rule.
                     * @returns {GC.Spread.Sheets.Style}
                     */
                    getExpected(): GC.Spread.Sheets.Style;
                    /**
                     * Specifies whether the range for this rule intersects another range.
                     * @param {number} row The row index.
                     * @param {number} column The column index.
                     * @param {number} rowCount The number of rows.
                     * @param {number} columnCount The number of columns.
                     * @returns {boolean} <c>true</c> if the range for this rule intersects another range; otherwise, <c>false</c>.
                     */
                    intersects(row:  number,  column:  number,  rowCount:  number,  columnCount:  number): boolean;
                    /**
                     * Specifies whether this rule is a scale rule.
                     * @returns {boolean} <c>true</c> if this rule is a scale rule; otherwise, <c>false</c>.
                     */
                    isScaleRule(): boolean;
                    /**
                     * Gets or sets the priority of the rule.
                     * @param {number} value The priority of the rule.
                     * @returns {number | GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase} If no value is set, returns the priority of the rule; otherwise, returns the condition rule.
                     */
                    priority(value?:  number): any;
                    /**
                     * Gets or sets the condition rule ranges.
                     * @param {Array.<GC.Spread.Sheets.Range>} value The condition rule ranges.
                     * @returns {Array.<GC.Spread.Sheets.Range> | GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase} If no value is set, returns the condition rule ranges; otherwise, returns the condition rule.
                     */
                    ranges(value?:  GC.Spread.Sheets.Range[]): any;
                    /**
                     * Resets the rule.
                     */
                    reset(): void;
                    /**
                     * Gets or sets the condition rule type.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.RuleType} value The condition rule type.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.RuleType | GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase} If no value is set, returns the condition rule type; otherwise, returns the condition rule.
                     */
                    ruleType(value?:  RuleType): any;
                    /**
                     * Gets or sets whether rules with lower priority are applied before this rule.
                     * @param {boolean} value Whether rules with lower priority are applied before this rule.
                     * @returns {boolean | GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase} If no value is set, returns whether the rules with lower priority are not applied before this rule; otherwise, returns the condition rule.
                     */
                    stopIfTrue(value?:  boolean): any;
                    /**
                     * Gets or sets the style for the rule.
                     * @param {GC.Spread.Sheets.Style} value The style for the rule.
                     * @returns {GC.Spread.Sheets.Style | GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase} If no value is set, returns the style for the rule; otherwise, returns the condition rule.
                     */
                    style(value?:  GC.Spread.Sheets.Style): any;
                }

                export class DataBarRule extends ConditionRuleBase{
                    /**
                     * Represents a data bar conditional rule with the specified parameters.
                     * @extends GC.Spread.Sheets.ConditionalFormatting.ScaleRule
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} minType The minimum scale type.
                     * @param {object} minValue The minimum scale value.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} maxType The maximum scale type.
                     * @param {object} maxValue The maximum scale value.
                     * @param {string} color The fill color of the data bar.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The data bar rule effected range.
                     * @class
                     */
                    constructor(minType:  ScaleValueType,  minValue:  Object,  maxType:  ScaleValueType,  maxValue:  Object,  color:  string,  ranges:  GC.Spread.Sheets.Range[]);
                    /**
                     * Gets or sets the axis color of the data bar.
                     * @param {string} value The axis color of the data bar.
                     * @returns {string | GC.Spread.Sheets.ConditionalFormatting.DataBarRule} If no value is set, returns the axis color of the data bar; otherwise, returns the data bar rule.
                     */
                    axisColor(value?:  string): any;
                    /**
                     * Gets or sets the axis position of the data bar.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.DataBarAxisPosition} value The axis position of the data bar.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.DataBarAxisPosition | GC.Spread.Sheets.ConditionalFormatting.DataBarRule} If no value is set, returns the axis position of the data bar; otherwise, returns the data bar rule.
                     */
                    axisPosition(value?:  DataBarAxisPosition): any;
                    /**
                     * Gets or sets the color of the border.
                     * @param {string} value The color of the border.
                     * @returns {string | GC.Spread.Sheets.ConditionalFormatting.DataBarRule} If no value is set, returns the color of the border; otherwise, returns the data bar rule.
                     */
                    borderColor(value?:  string): any;
                    /**
                     * Gets or sets the postive fill color of the data bar.
                     * @param {string} value The fill color.
                     * @returns {string | GC.Spread.Sheets.ConditionalFormatting.DataBarRule} If no value is set, returns the postive fill color of the data bar; otherwise, returns the data bar rule.
                     */
                    color(value?:  string): any;
                    /**
                     * Gets or sets the data bar direction.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.BarDirection} value The data bar direction.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.BarDirection | GC.Spread.Sheets.ConditionalFormatting.DataBarRule} If no value is set, returns the data bar direction; otherwise, returns the data bar rule.
                     */
                    dataBarDirection(value?:  BarDirection): any;
                    /**
                     * Returns the specified value of the rule if the cell meets the condition.
                     * @param {object} evaluator The evaluator.
                     * @param {number} baseRow The row index.
                     * @param {number} baseColumn The column index.
                     * @param {object} actual The current value.
                     * @returns {object} The specified value of the rule if the cell meets the condition.
                     */
                    evaluate(evaluator:  Object,  baseRow:  number,  baseColumn:  number,  actual:  Object): Object;
                    /**
                     * Gets or sets a value that indicates whether the data bar is a gradient.
                     * @param {boolean} value Whether the data bar is a gradient.
                     * @returns {boolean | GC.Spread.Sheets.ConditionalFormatting.DataBarRule} If no value is set, returns the value that indicates whether the data bar is a gradient; otherwise, returns the data bar rule.
                     */
                    gradient(value?:  boolean): any;
                    /**
                     * Gets or sets the color of the negative border.
                     * @param {string} value The color of the negative boreder.
                     * @returns {string | GC.Spread.Sheets.ConditionalFormatting.DataBarRule} If no value is set, returns the color of the negative border; otherwise, returns the data bar rule.
                     */
                    negativeBorderColor(value?:  string): any;
                    /**
                     * Gets or sets the color of the negative fill.
                     * @param {string} value The color of the negative fill.
                     * @returns {string | GC.Spread.Sheets.ConditionalFormatting.DataBarRule} If no value is set, returns the color of the negative fill; otherwise, returns the data bar rule.
                     */
                    negativeFillColor(value?:  string): any;
                    /**
                     * Gets or sets whether to display the data bar without text.
                     * @param {boolean} value Whether to display the data bar without text.
                     * @returns {boolean | GC.Spread.Sheets.ConditionalFormatting.DataBarRule} If no value is set, returns whether the widget displays the data bar without text; otherwise, returns the data bar rule.
                     */
                    showBarOnly(value?:  boolean): any;
                    /**
                     * Gets or sets a value that indicates whether to paint the border.
                     * @param {boolean} value Whether to paint the border.
                     * @returns {boolean | GC.Spread.Sheets.ConditionalFormatting.DataBarRule} If no value is set, returns the value that indicates whether to paint the border; otherwise, returns the data bar rule.
                     */
                    showBorder(value?:  boolean): any;
                    /**
                     * Gets or sets a value that indicates whether the negative border color is used to paint the border for the negative value.
                     * @param {boolean} value Whether the negative border color is used to paint the border for the negative value.
                     * @returns {boolean | GC.Spread.Sheets.ConditionalFormatting.DataBarRule} If no value is set, returns the value that indicates whether the negative border color is used to paint the border for the negative value; otherwise, returns the data bar rule.
                     */
                    useNegativeBorderColor(value?:  boolean): any;
                    /**
                     * Gets or sets a value that indicates whether the negative fill color is used to paint the negative value.
                     * @param {boolean} value Whether the negative fill color is used to paint the negative value.
                     * @returns {boolean | GC.Spread.Sheets.ConditionalFormatting.DataBarRule} If no value is set, returns the value that indicates whether the negative fill color is used to paint the negative value; otherwise, returns the data bar rule.
                     */
                    useNegativeFillColor(value?:  boolean): any;
                }

                export class IconCriterion{
                    /**
                     * Represents an icon criteria with the specified parameters.
                     * @class
                     * @param {boolean} isGreaterThanOrEqualTo If set to true, use the greater than or equal to operator to calculate the value.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.IconValueType} iconValueType The type of scale value.
                     * @param {object} iconValue The scale value.
                     */
                    constructor(isGreaterThanOrEqualTo:  boolean,  iconValueType:  IconValueType,  iconValue:  Object);
                }

                export class IconSetRule extends ConditionRuleBase{
                    /**
                     * Represents an icon set rule with the specified parameters.
                     * @class
                     * @extends GC.Spread.Sheets.ConditionalFormatting.ScaleRule
                     * @param {GC.Spread.Sheets.ConditionalFormatting.IconSetType} iconSetType The type of icon set.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges
                     */
                    constructor(iconSetType:  IconSetType,  ranges:  GC.Spread.Sheets.Range[]);
                    /**
                     * Returns the specified value of the rule if the cell meets the condition.
                     * @param {object} evaluator The evaluator.
                     * @param {number} baseRow The row index.
                     * @param {number} baseColumn The column index.
                     * @param {object} actual The current value.
                     * @returns {object} The specified value of the rule if the cell meets the condition.
                     */
                    evaluate(evaluator:  Object,  baseRow:  number,  baseColumn:  number,  actual:  Object): Object;
                    /**
                     * Gets the icon based on the specific iconSetType and iconIndex objects.
                     * @static
                     * @param {GC.Spread.Sheets.ConditionalFormatting.IconSetType} iconSetType The icon set type.
                     * @param {number} iconIndex The icon index.
                     * returns {object} An object that contains the image's URL string, the offset, and width and height.
                     * If the user wants to customize the icon for IconSet, it returns the image URL string.
                     */
                    static getIcon(iconSetType:  IconSetType,  iconIndex:  number): Object;
                    /**
                     * Gets the icon criteria.
                     * @returns {Array.<GC.Spread.Sheets.ConditionalFormatting.IconCriterion>} Returns the icon criterias whose item type is GC.Spread.Sheets.ConditionalFormatting.IconCriterion.
                     */
                    iconCriteria(): IconCriterion[];
                    /**
                     * Gets or sets the type of icon set.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.IconSetType} value The type of icon set.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.IconSetType | GC.Spread.Sheets.ConditionalFormatting.IconSetRule} If no value is set, returns the type of icon set; otherwise, returns the icon set rule.
                     */
                    iconSetType(value?:  IconSetType): any;
                    /**
                     * Resets the rule.
                     */
                    reset(): void;
                    /**
                     * Gets or sets whether to reverse icon order.
                     * @param {boolean} value Whether to reverse icon order.
                     * @returns {boolean | GC.Spread.Sheets.ConditionalFormatting.IconSetRule} If no value is set, returns the value that indicates whether to reverse icon order; otherwise, returns the icon set rule.
                     */
                    reverseIconOrder(value?:  boolean): any;
                    /**
                     * Gets or sets whether to display the icon only.
                     * @param {boolean} value Whether to display the icon only.
                     * @returns {boolean | GC.Spread.Sheets.ConditionalFormatting.IconSetRule} If no value is set, returns the value that indicates whether to display the icon only; otherwise, returns the icon set rule.
                     */
                    showIconOnly(value?:  boolean): any;
                }

                export class NormalConditionRule extends ConditionRuleBase{
                    /**
                     * Represents a normal conditional rule.
                     * @class
                     * @extends GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase
                     * @param {GC.Spread.Sheets.ConditionalFormatting.RuleType} ruleType
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The cell ranges where the rule is applied whose item type is GC.Spread.Sheets.Range.
                     * @param {GC.Spread.Sheets.Style} style The style that is applied to the cell when the condition is met.
                     * @param {object}operator The comparison operator.
                     * @param {object} value1 The first value.
                     * @param {object} value2 The second value.
                     * @param {string} text The text for comparison.
                     * @param {string} formula The condition formula.
                     * @param {object} type The average condition type.
                     * @param {number} rank The number of top or bottom items to apply the style to.
                     * @constructor
                     */
                    constructor(ruleType:  RuleType,  ranges:  GC.Spread.Sheets.Range[],  style:  GC.Spread.Sheets.Style,  operator:  any,  value1:  Object,  value2:  Object,  text:  string,  formula:  string,  type:  any,  rank:  number);
                    /**
                     * Creates a condition for the rule.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.Condition} The condition.
                     */
                    createCondition(): Condition;
                    /**
                     * Gets or sets the condition formula.
                     * @param {string} value The condition formula.
                     * @returns {string | GC.Spread.Sheets.ConditionalFormatting.NormalConditionRule} If no value is set, returns the condition formula; otherwise, returns the number condition rule.
                     */
                    formula(value?:  string): any;
                    /**
                     * Gets or sets the comparison operator.
                     * @param {object} value The comparison operator.
                     * @returns {object | GC.Spread.Sheets.ConditionalFormatting.NormalConditionRule} If no value is set, returns the comparison operator; otherwise, returns the number condition rule.
                     */
                    operator(value?:  Object): any;
                    /**
                     * Gets or sets the number of top or bottom items to apply the style to.
                     * @param {number} value The number of top or bottom items to apply the style to.
                     * @returns {number | GC.Spread.Sheets.ConditionalFormatting.NormalConditionRule} If no value is set, returns the number of top or bottom items to apply the style to; otherwise, returns the number condition rule.
                     */
                    rank(value?:  number): any;
                    /**
                     * Resets the rule.
                     */
                    reset(): void;
                    /**
                     * Gets or sets the text for comparison.
                     * @param {string} value The text for comparison.
                     * @returns {string | GC.Spread.Sheets.ConditionalFormatting.NormalConditionRule} If no value is set, returns the text for comparison; otherwise, returns the number condition rule.
                     */
                    text(value?:  string): any;
                    /**
                     * Gets or sets the average condition type.
                     * @param {object} value The average condition type.
                     * @returns {object | GC.Spread.Sheets.ConditionalFormatting.NormalConditionRule} If no value is set, returns the average condition type; otherwise, returns the number condition rule.
                     */
                    type(value?:  Object): any;
                    /**
                     * Gets or sets the first value.
                     * @param {object} value The first value.
                     * @returns {object | GC.Spread.Sheets.ConditionalFormatting.NormalConditionRule} If no value is set, returns the first value; otherwise, returns the number condition rule.
                     */
                    value1(value?:  Object): any;
                    /**
                     * Gets or sets the second value.
                     * @param {object} value The second value.
                     * @returns {object | GC.Spread.Sheets.ConditionalFormatting.NormalConditionRule} If no value is set, returns the second value; otherwise, returns the number condition rule.
                     */
                    value2(value?:  Object): any;
                }

                export class ScaleRule extends ConditionRuleBase{
                    /**
                     * Represents a scale conditional rule.
                     * @@extends GC.Spread.Sheets.ConditionalFormatting.ConditionRuleBase
                     * @param {GC.Spread.Sheets.ConditionalFormatting.RuleType} ruleType The rule type.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} minType The minimum scale type.
                     * @param minValue The minimum scale value.
                     * @param {string} minColor The minimum scale color.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} midType The midpoint scale type.
                     * @param midValue The midpoint scale value.
                     * @param {string} midColor The midpoint scale color.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} maxType The maximum scale type.
                     * @param maxValue The maximum scale value.
                     * @param {string} maxColor The maximum scale color.
                     * @param {Array.<GC.Spread.Sheets.Range>} ranges The ranges.
                     * @constructor
                     */
                    constructor(ruleType:  RuleType,  minType:  ScaleValueType,  minValue:  Object,  minColor:  string,  midType:  ScaleValueType,  midValue:  Object,  midColor:  string,  maxType:  ScaleValueType,  maxValue:  Object,  maxColor:  string,  ranges:  GC.Spread.Sheets.Range[]);
                    /**
                     * Creates a condition for the rule.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.Condition} The condition.
                     */
                    createCondition(): Condition;
                    /**
                     * Returns a specified value of the rule if the cell satisfies the condition.
                     * @param {object} evaluator The evaluator.
                     * @param {number} baseRow The row index.
                     * @param {number} baseColumn The column index.
                     * @param {object} actual The actual value object for evaluation.
                     * @returns {string} A specified value of the rule if the cell satisfies the condition.
                     */
                    evaluate(evaluator:  Object,  baseRow:  number,  baseColumn:  number,  actual:  Object): Object;
                    /**
                     * Gets or sets the maximum color scale.
                     * @param {string} value The maximum color scale.
                     * @returns {string | GC.Spread.Sheets.ConditionalFormatting.ScaleRule} If no value is set, returns the maximum color scale; otherwise, returns the scale rule.
                     */
                    maxColor(value?:  string): any;
                    /**
                     * Gets or sets the maximum scale type.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} value The maximum scale type.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType | GC.Spread.Sheets.ConditionalFormatting.ScaleRule} If no value is set, returns the maximum scale type; otherwise, returns the scale rule.
                     */
                    maxType(value?:  ScaleValueType): any;
                    /**
                     * Gets or sets the maximum scale value.
                     * @param {object} value The maximum scale value.
                     * @returns {object | GC.Spread.Sheets.ConditionalFormatting.ScaleRule} If no value is set, returns the maximum scale value; otherwise, returns the scale rule.
                     */
                    maxValue(value?:  Object): any;
                    /**
                     * Gets or sets the midpoint scale color.
                     * @param {string} value The midpoint scale color.
                     * @returns {string | GC.Spread.Sheets.ConditionalFormatting.ScaleRule} If no value is set, returns the midpoint scale color; otherwise, returns the scale rule.
                     */
                    midColor(value?:  string): any;
                    /**
                     * Gets or sets the midpoint scale type.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} value The midpoint scale type.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType | GC.Spread.Sheets.ConditionalFormatting.ScaleRule} If no value is set, returns the midpoint scale type; otherwise, returns the scale rule.
                     */
                    midType(value?:  ScaleValueType): any;
                    /**
                     * Gets or sets the midpoint scale value.
                     * @param {object} value The midpoint scale value.
                     * @returns {object | GC.Spread.Sheets.ConditionalFormatting.ScaleRule} If no value is set, returns the midpoint scale value; otherwise, returns the scale rule.
                     */
                    midValue(value?:  Object): any;
                    /**
                     * Gets or sets the minimum scale color.
                     * @param {string} value The minimum scale color.
                     * @returns {string | GC.Spread.Sheets.ConditionalFormatting.ScaleRule} If no value is set, returns the minimum scale color; otherwise, returns the scale rule.
                     */
                    minColor(value?:  string): any;
                    /**
                     * Gets or sets the type of minimum scale.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType} value The type of minimum scale.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.ScaleValueType | GC.Spread.Sheets.ConditionalFormatting.ScaleRule} If no value is set, returns the type of minimum scale; otherwise, returns the scale rule.
                     */
                    minType(value?:  ScaleValueType): any;
                    /**
                     * Gets or sets the minimum scale value.
                     * @param {object} value The minimum scale value.
                     * @returns {object | GC.Spread.Sheets.ConditionalFormatting.ScaleRule} If no value is set, returns the minimum scale value; otherwise, returns the scale rule.
                     */
                    minValue(value?:  Object): any;
                    /**
                     * Gets whether evaluation should stop if the condition evaluates to <c>true</c>.
                     */
                    stopIfTrue(value:  boolean): boolean;
                }

                export class ScaleValue{
                    /**
                     * Represents a scale value with the specified type and value.
                     * @class
                     * @param {object} type The scale value type.
                     * @param {object} value The scale value.
                     */
                    constructor(type:  Object,  value:  Object);
                    /** Gets the scale value type.
                     * @type {object}
                     */
                    type: Object;
                    /** Gets the scale value.
                     * @type {object}
                     */
                    value: Object;
                }
            }

            module DataValidation{
                /**
                 * Creates a validator based on the data.
                 * @static
                 * @param {GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators} typeOperator The type of ComparisonOperators compare operator.
                 * @param {object} v1 The first object.
                 * @param {object} v2 The second object.
                 * @returns {GC.Spread.Sheets.DataValidation.DefaultDataValidator} The validator.
                 */
                function createDateValidator(typeOperator:  GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators,  v1:  Object,  v2:  Object): GC.Spread.Sheets.DataValidation.DefaultDataValidator;
                /**
                 * Creates a validator based on a formula list.
                 * @static
                 * @param {string} formula The formula list.
                 * @returns {GC.Spread.Sheets.DataValidation.DefaultDataValidator} The validator.
                 */
                function createFormulaListValidator(formula:  string): GC.Spread.Sheets.DataValidation.DefaultDataValidator;
                /**
                 * Creates a validator based on a formula.
                 * @static
                 * @param {string} formula The formula condition.
                 * @returns {GC.Spread.Sheets.DataValidation.DefaultDataValidator} The validator.
                 */
                function createFormulaValidator(formula:  string): GC.Spread.Sheets.DataValidation.DefaultDataValidator;
                /**
                 * Creates a validator based on a list.
                 * @static
                 * @param {string} source The list value.
                 * @returns {GC.Spread.Sheets.DataValidation.DefaultDataValidator} The validator.
                 */
                function createListValidator(source:  string): GC.Spread.Sheets.DataValidation.DefaultDataValidator;
                /**
                 * Creates a validator based on numbers.
                 * @static
                 * @param {GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators} typeOperator The type of ComparisonOperators compare operator.
                 * @param {object} v1 The first object.
                 * @param {object} v2 The second object.
                 * @param {boolean} isIntegerValue Set to <c>true</c> if the validator is set to a number.
                 * @returns {GC.Spread.Sheets.DataValidation.DefaultDataValidator} The validator.
                 */
                function createNumberValidator(typeOperator:  GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators,  v1:  Object,  v2:  Object,  isIntegerValue:  boolean): GC.Spread.Sheets.DataValidation.DefaultDataValidator;
                /**
                 * Creates a validator based on text length.
                 * @static
                 * @param {GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators} typeOperator The type of ComparisonOperators compare operator.
                 * @param {object} v1 The first object.
                 * @param {object} v2 The second object.
                 * @returns {GC.Spread.Sheets.DataValidation.DefaultDataValidator} The validator.
                 */
                function createTextLengthValidator(typeOperator:  GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators,  v1:  Object,  v2:  Object): GC.Spread.Sheets.DataValidation.DefaultDataValidator;
                /**
                 * Indicates the data validator criteria type.
                 * @enum {number}
                 */
                export enum CriteriaType{
                    /**
                     * Specifies that the data validation allows any type of value and does not check for a type or range of values.
                     */
                    anyValue= 0,
                    /**
                     * Specifies that the data validation checks for and allows whole number values satisfying the given condition.
                     */
                    wholeNumber= 0x1,
                    /**
                     * Specifies that the data validation checks for and allows decimal values satisfying the given condition.
                     */
                    decimalValues= 0x2,
                    /**
                     * Specifies that the data validation checks for and allows a value that matches one in a list of values.
                     */
                    list= 0x3,
                    /**
                     * Specifies that the data validation checks for and allows date values satisfying the given condition.
                     */
                    date= 0x4,
                    /**
                     * Specifies that the data validation checks for and allows time values satisfying the given condition.
                     */
                    time= 0x5,
                    /**
                     * Specifies that the data validation checks for and allows text values whose length satisfies the given condition.
                     */
                    textLength= 0x6,
                    /**
                     * Specifies that the data validation uses a custom formula to check the cell value.
                     */
                    custom= 0x7
                }

                /**
                 * Indicates the data validation result.
                 * @enum {number}
                 */
                export enum DataValidationResult{
                    /**
                     * Indicates to apply the value to a cell for a validation error.
                     */
                    forceApply= 0,
                    /**
                     * Indicates to discard the value and not apply it to the cell for a validation error.
                     */
                    discard= 1,
                    /**
                     * Indicates to retry multiple times to apply the value to the cell for a validation error.
                     */
                    retry= 2
                }

                /**
                 * Indicates the data validation error style.
                 * @enum {number}
                 */
                export enum ErrorStyle{
                    /**
                     * Specifies to use a stop icon in the error alert.
                     */
                    stop= 0,
                    /**
                     * Specifies to use a warning icon in the error alert.
                     */
                    warning= 1,
                    /**
                     * Specifies to use an information icon in the error alert.
                     */
                    information= 2
                }


                export class DefaultDataValidator{
                    /**
                     * Represents a data validator.
                     * @class
                     * @param {GC.Spread.Sheets.ConditionalFormatting.Condition} condition The condition.
                     */
                    constructor(condition:  GC.Spread.Sheets.ConditionalFormatting.Condition);
                    /**
                     * Gets or sets the comparison operator.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators} value The comparison operator.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators | GC.Spread.Sheets.DataValidation.DefaultDataValidator} If no value is set, returns the comparison operator; otherwise, returns the data validator.
                     */
                    comparisonOperator(value?:  GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators): any;
                    /**
                     * Gets or sets the condition to validate.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.Condition} value The condition to validate.
                     * @returns {GC.Spread.Sheets.ConditionalFormatting.Condition | GC.Spread.Sheets.DataValidation.DefaultDataValidator} If no value is set, returns the condition to validate; otherwise, returns the data validator.
                     */
                    condition(value?:  GC.Spread.Sheets.ConditionalFormatting.Condition): any;
                    /**
                     * Gets or sets the error message.
                     * @param {string} value The error message.
                     * @returns {string | GC.Spread.Sheets.DataValidation.DefaultDataValidator} If no value is set, returns the error message; otherwise, returns the data validator.
                     */
                    errorMessage(value?:  string): any;
                    /**
                     * Gets or sets the error style to display.
                     * @param {GC.Spread.Sheets.DataValidation.ErrorStyle} value The error style to display.
                     * @returns {GC.Spread.Sheets.DataValidation.ErrorStyle | GC.Spread.Sheets.DataValidation.DefaultDataValidator} If no value is set, returns the error style to display; otherwise, returns the data validator.
                     */
                    errorStyle(value?:  GC.Spread.Sheets.DataValidation.ErrorStyle): any;
                    /**
                     * Gets or sets the error title.
                     * @param {string} value The error title.
                     * @returns {string | GC.Spread.Sheets.DataValidation.DefaultDataValidator} If no value is set, returns the error title; otherwise, returns the data validator.
                     */
                    errorTitle(value?:  string): any;
                    /**
                     * Returns the valid data lists if the Data validation type is list; otherwise, returns null.
                     * @param {object} evaluator The object that can evaluate a condition.
                     * @param {number} baseRow The base row.
                     * @param {number} baseColumn The base column.
                     * @returns {Array} The valid data lists or null.
                     */
                    getValidList(evaluator:  Object,  baseRow:  number,  baseColumn:  number): any[];
                    /**
                     * Gets or sets whether to ignore an empty value.
                     * @param {boolean} value Indicates whether to ignore the empty value.
                     * @returns {boolean | GC.Spread.Sheets.DataValidation.DefaultDataValidator} If no value is set, returns whether to ignore the empty value; otherwise, returns the data validator.
                     */
                    ignoreBlank(value?:  boolean): any;
                    /**
                     * Gets or sets whether to display a drop-down button.
                     * @param {boolean} value Indicates whether to display a drop-down button.
                     * @returns {boolean | GC.Spread.Sheets.DataValidation.DefaultDataValidator} If no value is set, returns whether to display a drop-down button; otherwise, returns the data validator.
                     */
                    inCellDropdown(value?:  boolean): any;
                    /**
                     * Gets or sets the input message.
                     * @param {string} value The input message.
                     * @returns {string | GC.Spread.Sheets.DataValidation.DefaultDataValidator} If no value is set, returns the input message; otherwise, returns the data validator.
                     */
                    inputMessage(value?:  string): any;
                    /**
                     * Gets or sets the input title.
                     * @param {string} value The input title.
                     * @returns {string | GC.Spread.Sheets.DataValidation.DefaultDataValidator} If no value is set, returns the input title; otherwise, returns the data validator.
                     */
                    inputTitle(value?:  string): any;
                    /**
                     * Determines whether the current value is valid.
                     * @param {object} evaluator The evaluator.
                     * @param {number} baseRow The base row.
                     * @param {number} baseColumn The base column.
                     * @param {object} actual The current value.
                     * @returns {boolean} <c>true</c> if the value is valid; otherwise, <c>false</c>.
                     */
                    isValid(evaluator:  Object,  baseRow:  number,  baseColumn:  number,  actual:  Object): boolean;
                    /**
                     * Resets the data validator.
                     */
                    reset(): void;
                    /**
                     * Gets or sets whether to display an error message.
                     * @param {boolean} value Indicates whether to display an error message.
                     * @returns {boolean | GC.Spread.Sheets.DataValidation.DefaultDataValidator} If no value is set, returns whether to display an error message; otherwise, returns the data validator.
                     */
                    showErrorMessage(value?:  boolean): any;
                    /**
                     * Gets or sets whether to display the input title and input message.
                     * @param {boolean} value Indicates whether to display the input title and input message.
                     * @returns {boolean | GC.Spread.Sheets.DataValidation.DefaultDataValidator} If no value is set, returns whether to display the input title and input message; otherwise, returns the data validator.
                     */
                    showInputMessage(value?:  boolean): any;
                    /**
                     * Gets or sets the criteria type of this data validator.
                     * @param {GC.Spread.Sheets.DataValidation.CriteriaType} value The criteria type of this data validator.
                     * @returns {GC.Spread.Sheets.DataValidation.CriteriaType | GC.Spread.Sheets.DataValidation.DefaultDataValidator} If no value is set, returns the criteria type of this data validator; otherwise, returns the data validator.
                     */
                    type(value?:  GC.Spread.Sheets.DataValidation.CriteriaType): any;
                    /**
                     * Gets the first value of the data validation.
                     * @returns {object} The first value.
                     */
                    value1(): any;
                    /**
                     * Gets the second value of the data validation.
                     * @returns {object} The second value.
                     */
                    value2(): any;
                }
            }

            module Fill{
                // <editor-fold desc="AutoFillType">
                 /**
                 * Represents the type of drag fill.
                 * @enum {number}
                 */
                export enum AutoFillType{
                    /**
                     *  Fills cells with all data objects, including values, formatting, and formulas.
                     */
                    copyCells= 0,
                    /**
                     *   Fills cells with series.
                     */
                    fillSeries= 1,
                    /**
                     *   Fills cells only with formatting.
                     */
                    fillFormattingOnly= 2,
                    /**
                     *   Fills cells with values and not formatting.
                     */
                    fillWithoutFormatting= 3,
                    /**
                     *  Clears cell values.
                     */
                    clearValues= 4,
                    /**
                     *  Automatically fills cells.
                     */
                    auto= 5
                }

                // <editor-fold desc="FillDateUnit">
                 /**
                 * Represents the date fill unit.
                 * @enum {number}
                 */
                export enum FillDateUnit{
                    /** Sets the date fill unit to day.
                     * @type {number}
                     */
                    day= 0,
                    /** Sets the date fill unit to weekday.
                     * @type {number}
                     */
                    weekday= 1,
                    /** Sets the date fill unit to month.
                     * @type {number}
                     */
                    month= 2,
                    /** Sets the date fill unit to year.
                     * @type {number}
                     */
                    year= 3
                }

                // <editor-fold desc="FillDirection">
                 /**
                 * Represents the type of drag fill direction.
                 * @enum {number}
                 */
                export enum FillDirection{
                    /**
                     *  Fills from the right to the left.
                     */
                    left= 0,
                    /**
                     *  Fills from the left to the right.
                     */
                    right= 1,
                    /**
                     *  Fills from the bottom to the top.
                     */
                    up= 2,
                    /**
                     *   Fills from the top to the bottom.
                     */
                    down= 3
                }

                // <editor-fold desc="FillSeries">
                 /**
                 * Represents the fill series for drag fill.
                 * @enum {number}
                 */
                export enum FillSeries{
                    /**
                     *  Fills the column data.
                     */
                    column= 0,
                    /**
                     *  Fills the row data.
                     */
                    row= 1
                }

                // <editor-fold desc="FillType">
                 /**
                 * Represents the type of fill data.
                 * @enum {number}
                 */
                export enum FillType{
                    /** Represents the direction fill type.
                     * @type {number}
                     */
                    direction= 0,
                    /** Represents the linear fill type.
                     * @type {number}
                     */
                    linear= 1,
                    /** Represents the growth fill type.
                     * @type {number}
                     */
                    growth= 2,
                    /** Represents the date fill type.
                     * @type {number}
                     */
                    date= 3,
                    /** Represents the auto fill type.
                     * @type {number}
                     */
                    auto= 4
                }

            }

            module Filter{

                export interface IFilteredArgs{
                    action: FilterActionType;
                    sheet: Sheets.Worksheet;
                    range: Sheets.Range;
                    filteredRows: number[];
                    filteredOutRows: number[];
                    columns: number;
                }

                /**
                 * Defines the type of filter action.
                 * @enum {number}
                 */
                export enum FilterActionType{
                    /** Specifies the filter action.
                     */
                    filter= 0,
                    /** Specifies the unfilter action.
                     */
                    unfilter= 1
                }


                export class HideRowFilter extends RowFilterBase{
                    /**
                     * Represents a default row filter.
                     * @class GC.Spread.Sheets.Filter.HideRowFilter
                     * @extends GC.Spread.Sheets.Filter.RowFilterBase
                     * @param {GC.Spread.Sheets.Range} range The filter range.
                     */
                    constructor(range?:  GC.Spread.Sheets.Range);
                }

                export class RowFilterBase{
                    /**
                     * Represents a row filter base that supports row filters for filtering rows in a sheet.
                     * @class GC.Spread.Sheets.Filter.RowFilterBase
                     * @param {GC.Spread.Sheets.Range} range The filter range.
                     */
                    constructor(range?:  GC.Spread.Sheets.Range);
                    /**
                     * Represents the range for the row filter.
                     * @type {GC.Spread.Sheets.Range}
                     */
                    range: GC.Spread.Sheets.Range;
                    /**
                     * Represents the type name string used for supporting serialization.
                     * @type {string}
                     */
                    typeName: string;
                    /**
                     * Adds a specified filter to the row filter.
                     * @param {number} col The column index.
                     * @param {GC.Spread.Sheets.ConditionalFormatting.Condition} condition The condition to filter.
                     */
                    addFilterItem(col:  number,  condition:  GC.Spread.Sheets.ConditionalFormatting.Condition): void;
                    /**
                     * Filters the specified column.
                     * @param {number} col The index of the column to be filtered; if it is omitted, all the columns in the range will be filtered.
                     */
                    filter(col?:  number): void;
                    /**
                     * Gets or sets whether the sheet column's filter button is displayed.
                     * @param {number} col The column index of the filter button.
                     * @param {boolean} value Whether the filter button is displayed.
                     * @returns {boolean | GC.Spread.Sheets.Filter.RowFilterBase}<br />
                     *      No parameter <c>false</c> if all filter buttons are invisible; otherwise, <c>true</c>.<br />
                     *      One parameter col <c>false</c> if the specified column filter button is invisible; otherwise, <c>true</c>.<br />
                     *      One parameter value <c>GC.Spread.Sheets.Filter.RowFilterBase</c> sets all filter buttons to be visible(true)/invisible(false).<br />
                     *      Two parameters col,value <c>GC.Spread.Sheets.Filter.RowFilterBase</c> sets the specified column filter button to be visible(true)/invisible(false).
                     */
                    filterButtonVisible(col?:  number,  value?:  boolean): any;
                    /**
                     * Loads the object state from the specified JSON string.
                     * @param {object} settings The row filter data from deserialization.
                     */
                    fromJSON(settings:  Object): void;
                    /**
                     * Gets all the filtered conditions.
                     * @returns {Array} Returns a collection that contains all the filtered conditions.
                     */
                    getFilteredItems(): any[];
                    /**
                     * Gets the filters for the specified column.
                     * @param {number} col The column index.
                     * @returns {Array} Returns a collection that contains conditions that belong to a specified column.
                     */
                    getFilterItems(col:  number): any[];
                    /**
                     * Gets the current sort state.
                     * @param {number} col The column index.
                     * @returns {GC.Spread.Sheets.SortState} The sort state of the current filter.
                     */
                    getSortState(col:  number): GC.Spread.Sheets.SortState;
                    /**
                     * Gets a value that indicates whether any row or specified column is filtered.
                     * @param {number} col The column index.
                     * @returns {boolean}  No parameter <c>true</c> if some rows are filtered; otherwise, <c>false</c>.<br />
                     *                     One parameter col <c>true</c> if the specified column is filtered; otherwise, <c>false</c>.
                     */
                    isFiltered(col?:  number): boolean;
                    /**
                     * Determines whether the specified row is filtered out.
                     * @param {number} row The row index.
                     * @returns {boolean} <c>true</c> if the row is filtered out; otherwise, <c>false</c>.
                     */
                    isRowFilteredOut(row:  number): boolean;
                    /**
                     * Performs the action when some columns have just been filtered or unfiltered.
                     * @param {object} args An object that contains the <i>action</i>, <i>sheet</i>, <i>range</i>, <i>filteredRows</i>, and <i>filteredOutRows</i>. See the Remarks for additional information.
                     */
                    onFilter(args:  IFilteredArgs): void;
                    /**
                     * Opens the filter dialog when the user clicks the filter button.
                     * @param {object} filterButtonHitInfo The hit test information about the filter button.
                     */
                    openFilterDialog(filterButtonHitInfo:  IFilterButtonHitInfo): void;
                    /**
                     * Removes the specified filter.
                     * @param {number} col The column index.
                     */
                    removeFilterItems(col:  number): void;
                    /**
                     * Clears all filters.
                     */
                    reset(): any;
                    /**
                     * Sorts the specified column in the specified order.
                     * @param {number} col The column index.
                     * @param {boolean} ascending Set to <c>true</c> to sort as ascending.
                     */
                    sortColumn(col:  number,  ascending:  boolean): void;
                    /**
                     * Saves the object state to a JSON string.
                     * @returns {object} The row filter data.
                     */
                    toJSON(): Object;
                    /**
                     * Removes the filter from the specified column.
                     * @param {number} col The index of the column for which to remove the filter; if it is omitted, removes the filter for all columns in the range.
                     */
                    unfilter(col?:  number): void;
                }
            }

            module FloatingObjects{

                export class FloatingObject{
                    /**
                     * Represents a floating object.
                     * @class
                     * @param {string} name The name of the floating object.
                     * @param {number} x The <i>x</i> location of the floating object.
                     * @param {number} y The <i>y</i> location of the floating object.
                     * @param {number} width The width of the floating object.
                     * @param {number} height The height of the floating object.
                     * @remarks
                     * This is a base class that is intended for internal use.
                     */
                    constructor(name:  string,  x:  number,  y:  number,  width:  number,  height:  number);
                    /** Represents the type name string used for supporting serialization.
                     * @type {string}
                     */
                    typeName: string;
                    /**
                     * Gets or sets whether to disable moving the floating object.
                     * @param {boolean} value The setting for whether to disable moving the floating object.
                     * @returns {boolean | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the setting for whether to disable moving the floating object; otherwise, returns the floating object.
                     */
                    allowMove(value?:  boolean): any;
                    /**
                     * Gets or sets whether to disable resizing the floating object.
                     * @param {boolean} value The setting for whether to disable resizing the floating object.
                     * @returns {boolean | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the setting for whether to disable resizing the floating object; otherwise, returns the floating object.
                     */
                    allowResize(value?:  boolean): any;
                    /**
                     * Gets a copy of the current content of the instance.
                     * @returns {HTMLElement} A copy of the current content of the instance.
                     */
                    cloneContent(): HTMLElement;
                    /**
                     * Gets or sets the content of the custom floating object.
                     * @param {HTMLElement} value The content of the custom floating object.
                     * @returns {object | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the content of the custom floating object; otherwise, returns the floating object.
                     */
                    content(value?:  HTMLElement): any;
                    /**
                     * Gets or sets whether the object moves when hiding or showing, resizing, or moving rows or columns.
                     * @param {boolean} value The value indicates whether the object moves when hiding or showing, resizing, or moving rows or columns.
                     * @returns {boolean | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns whether this floating object dynamically moves; otherwise, returns the floating object.
                     */
                    dynamicMove(value?:  boolean): any;
                    /**
                     * Gets or sets whether the size of the object changes when hiding or showing, resizing, or moving rows or columns.
                     * @param {boolean} value The value indicates whether the size of the object changes when hiding or showing, resizing, or moving rows or columns.
                     * @returns {boolean | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns whether this floating object dynamically changes size; otherwise, returns the floating object.
                     */
                    dynamicSize(value?:  boolean): any;
                    /**
                     * Gets or sets the end column index of the floating object position.
                     * @param {number} value The end column index of the floating object position.
                     * @returns {number | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the end column index of the floating object position; otherwise, returns the floating object.
                     */
                    endColumn(value?:  number): any;
                    /**
                     * Gets or sets the offset relative to the end column of the floating object.
                     * @param {number} value The offset relative to the end column of the floating object.
                     * @returns {number | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the offset relative to the end column of the floating object; otherwise, returns the floating object.
                     */
                    endColumnOffset(value?:  number): any;
                    /**
                     * Gets or sets the end row index of the floating object position.
                     * @param {number} value The end row index of the floating object position.
                     * @returns {number | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the end row index of the floating object position; otherwise, returns the floating object.
                     */
                    endRow(value?:  number): any;
                    /**
                     * Gets or sets the offset relative to the end row of the floating object.
                     * @param {number} value The offset relative to the end row of the floating object.
                     * @returns {number | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the offset relative to the end row of the floating object; otherwise, returns the floating object.
                     */
                    endRowOffset(value?:  number): any;
                    /**
                     * Gets or sets whether the position of the floating object is fixed. When fixedPosition is true, dynamicMove and dynamicSize are disabled.
                     * @param {boolean} value The value indicates whether the position of the floating object is fixed.
                     * @returns {boolean | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns whether the position of the floating object is fixed; otherwise, returns the floating object.
                     */
                    fixedPosition(value:  boolean): any;
                    /**
                     * Gets the dom host of the custom content.
                     * @returns {Array.<HTMLElement>}
                     */
                    getHost(): HTMLElement[];
                    /**
                     * Gets or sets the height of a floating object.
                     * @param {number} value The height of a floating object.
                     * @returns {number | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the height of a floating object; otherwise, returns the floating object.
                     */
                    height(value?:  number): any;
                    /**
                     * Gets or sets whether this floating object is locked.
                     * @param {boolean} value The value that indicates whether this floating object is locked.
                     * @returns {boolean | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns whether this floating object is locked; otherwise, returns the floating object.
                     */
                    isLocked(value?:  boolean): any;
                    /**
                     * Gets or sets whether this floating object is selected.
                     * @param {boolean} value The value that indicates whether this floating object is selected.
                     * @returns {boolean | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns whether this floating object is selected; otherwise, returns the floating object.
                     */
                    isSelected(value?:  boolean): any;
                    /**
                     * Gets or sets whether this floating object is visible.
                     * @param {boolean} value The value that indicates whether this floating object is visible.
                     * @returns {boolean | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns whether this floating object is visible; otherwise, returns the floating object.
                     */
                    isVisible(value?:  boolean): any;
                    /**
                     * Gets the name of the floating object.
                     * @param {string} value The name of the floating object.
                     * @returns {string | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the name of the floating object; otherwise, returns the floating object.
                     */
                    name(value?:  string): any;
                    /**
                     * Refresh the content in floatingObject.The user should override this method to make their content synchronize with the floatingObject.
                     */
                    refreshContent(): void;
                    /**
                     * Gets or sets the starting column index of the floating object position.
                     * @param {number} value The starting column index of the floating object position.
                     * @returns {number | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the starting column index of the floating object position; otherwise, returns the floating object.
                     */
                    startColumn(value?:  number): any;
                    /**
                     * Gets or sets the offset relative to the start column of the floating object.
                     * @param {number} value The offset relative to the start column of the floating object.
                     * @returns {number | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the offset relative to the start column of the floating object; otherwise, returns the floating object.
                     */
                    startColumnOffset(value?:  number): any;
                    /**
                     * Gets or sets the starting row index of the floating object position.
                     * @param {number} value The starting row index of the floating object position.
                     * @returns {number | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the starting row index of the floating object position; otherwise, returns the floating object.
                     */
                    startRow(value?:  number): any;
                    /**
                     * Gets or sets the offset relative to the start row of the floating object.
                     * @param {number} value The offset relative to the start row of the floating object.
                     * @returns {number | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the offset relative to the start row of the floating object; otherwise, returns the floating object.
                     */
                    startRowOffset(value?:  number): any;
                    /**
                     * Gets or sets the width of a floating object.
                     * @param {number} value The width of a floating object.
                     * @returns {number | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the width of a floating object; otherwise, returns the floating object.
                     */
                    width(value?:  number): any;
                    /**
                     * Gets or sets the horizontal location of the floating object.
                     * @param {number} value The horizontal location of the floating object.
                     * @return {number | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the horizontal location of the floating object; otherwise, returns the floating object.
                     */
                    x(value?:  number): any;
                    /**
                     * Gets or sets the vertical location of the floating object.
                     * @param {number} value The vertical location of the floating object.
                     * @return {number | GC.Spread.Sheets.FloatingObjects.FloatingObject} If no value is set, returns the vertical location of the floating object; otherwise, returns the floating object.
                     */
                    y(value?:  number): any;
                }

                export class FloatingObjectCollection{
                    /**
                     * Represents a floating object manager that managers all floating objects in a sheet.
                     * @class
                     * @param {GC.Spread.Sheets.Worksheet} sheet The worksheet.
                     * @param {string} typeName The type name.
                     */
                    constructor();
                    /**
                     * Adds a floating object to the sheet.<br />
                     * The arguments has 2 modes.<br />
                     * If there is 1 parameter, the parameter is floatingObject which is a GC.Spread.Sheets.FloatingObjects.FloatingObject type;
                     * If there are 6 parameters, the parameters are name, src, x, y, width, height.
                     * @param {GC.Spread.Sheets.FloatingObjects.FloatingObject|string} floatingObjectOrName The floating object that will be added to the sheet, or the name of the picture that will be added to the sheet.
                     * @param {string} src The image source of the picture.
                     * @param {number} x The x location of the picture.
                     * @param {number} y The y location of the picture.
                     * @param {number} width The width of the picture.
                     * @param {nubmer} height The height of the picture.
                     * @return {GC.Spread.Sheets.FloatingObjects.FloatingObject} The floating object that has been added to the sheet.
                     */
                    add(floatingObjectOrName:  Object,  src?:  string,  x?:  number,  y?:  number,  width?:  number,  height?:  number): Object;
                    /**
                     * Gets all of the floating objects in the sheet.
                     * @return {Array} The collection of all the floating objects in the sheet.
                     */
                    all(): GC.Spread.Sheets.FloatingObjects.FloatingObject[];
                    /**
                     * Removes all floating objects in the sheet.
                     */
                    clear(): void;
                    /**
                     * Gets a floating object from the sheet by the indicate name.
                     * @param {string} name The name of the floating object.
                     * @return {GC.Spread.Sheets.FloatingObjects.FloatingObject} The floating object in the sheet with the indicate name.
                     */
                    get(name:  string): GC.Spread.Sheets.FloatingObjects.FloatingObject;
                    /**
                     * Removes a floating object from the sheet by the indicate name.
                     * @param {string} name The name of the floating object.
                     */
                    remove(name:  string): GC.Spread.Sheets.FloatingObjects.FloatingObject;
                    /**
                     * Gets or sets the z-index of floating object.
                     * @param {string} name The name of the flaotingObject.
                     * @param {number} zIndex The z-index of the floating object.
                     * @return {number | *} If the parameter 'zIndex' is null or undefined,it will return the z-index of the floating object with the indicate name.
                     */
                    zIndex(name:  string,  zIndex?:  number): any;
                }

                export class Picture extends FloatingObject{
                    /**
                     * Represents a picture.
                     * @extends GC.Spread.Sheets.FloatingObjects.FloatingObject
                     * @class
                     * @param {string} name The name of the picture.
                     * @param {string} src The image source of the picture.
                     * @param {number} x The <i>x</i> location of the picture.
                     * @param {number} y The <i>y</i> location of the picture.
                     * @param {number} width The width of the picture.
                     * @param {number} height The height of the picture.
                     */
                    constructor(name:  string,  src:  string,  x:  number,  y:  number,  width:  number,  height:  number);
                    /**
                     * Gets or sets the background color of the picture.
                     * @param {string} value The backcolor of the picture.
                     * @returns {string | GC.Spread.Sheets.FloatingObjects.Picture} If no value is set, returns the backcolor of the picture; otherwise, returns the picture.
                     */
                    backColor(value?:  string): any;
                    /**
                     * Gets or sets the border color of the picture.
                     * @param {string} value The border color of the picture.
                     * @returns {Object | GC.Spread.Sheets.FloatingObjects.Picture} If no value is set, returns the border color of the picture; otherwise, returns the picture.
                     */
                    borderColor(value?:  string): any;
                    /**
                     * Gets or sets the border radius of the picture.
                     * @param {number} value The border radius of the picture.
                     * @returns {number | GC.Spread.Sheets.FloatingObjects.Picture} If no value is set, returns the border radius of the picture; otherwise, returns the picture.
                     */
                    borderRadius(value?:  number): any;
                    /**
                     * Gets or sets the border style of the picture.
                     * @param {string} value The css border style of the picture, such as dotted, dashed, solid, and so on.
                     * @returns {string | GC.Spread.Sheets.FloatingObjects.Picture} If no value is set, returns the border style of the picture; otherwise, returns the picture.
                     */
                    borderStyle(value?:  string): any;
                    /**
                     * Gets or sets the border width of the picture.
                     * @param {number} value The border width of the picture.
                     * @returns {number | GC.Spread.Sheets.FloatingObjects.Picture} If no value is set, returns the border width of the picture; otherwise, returns the picture.
                     */
                    borderWidth(value?:  number): any;
                    /**
                     * Gets the original height of the picture.
                     * @returns {number} The original height of the picture.
                     */
                    getOriginalHeight(): number;
                    /**
                     * Gets the original width of the picture.
                     * @returns {number} The original width of the picture.
                     */
                    getOriginalWidth(): number;
                    /**
                     * Gets or sets the stretch of the picture.
                     * @param {GC.Spread.Sheets.ImageLayout} value The stretch of the picture.
                     * @returns {GC.Spread.Sheets.ImageLayout | GC.Spread.Sheets.FloatingObjects.Picture} If no value is set, returns the stretch of the picture; otherwise, returns the picture.
                     */
                    pictureStretch(value?:  GC.Spread.Sheets.ImageLayout): any;
                    /**
                     * Gets or sets the src of the picture.
                     * @param {string} value The src of the picture.
                     * @returns {string | GC.Spread.Sheets.FloatingObjects.Picture} If no value is set, returns the src of the picture; otherwise, returns the picture.
                     */
                    src(value?:  string): any;
                }
            }

            module FormulaTextBox{

                export class FormulaTextBox{
                    /**
                     * Represents a formula text box.
                     * @class
                     * @param {Object} host The DOM element. It can be INPUT, TEXTAREA, or editable DIV.
                     * @param {Object} options The options. Default is {rangeSelectMode: false, absoluteReference: false}
                     */
                    constructor(host:  Object,  options:  Object);
                    /**
                     * Adds a custom function description.
                     * @param {IFunctionDescription} functionDescription The function description to add. This can be an array. See the Remarks for more information.
                     */
                    add(fnd:  Object): void;
                    /**
                     * Gets or sets whether the text box uses automatic complete.
                     * @param {boolean} value Whether to use automatic complete when editing.
                     * @returns {boolean} If no value is set, returns whether the text box uses auto complete; otherwise, there is no return value.
                     */
                    autoComplete(value?:  boolean): any;
                    /**
                     * Binds an event.
                     * @param {string} type The event type.
                     * @param {Object} data Optional. Specifies additional data to pass along to the function.
                     * @param {Function} fn Specifies the function to run when the event occurs.
                     */
                    bind(type:  string,  data:  Object,  fn:  Function): void;
                    /**
                     * Gets or sets the cursor position.
                     * @param {number} value The cursor position.
                     * @returns {number} If no value is set, returns the cursor position; otherwise, there is no return value.
                     */
                    caret(value?:  number): any;
                    /**
                     * Removes host from formula text box and removes all binding events.
                     */
                    destroy(): void;
                    /**
                     * refresh the formula text box with the active cell.
                     */
                    refresh(): void;
                    /**
                     * Removes a custom function description.
                     * @param {string} name The custom function description name.
                     */
                    remove(name:  string): void;
                    /**
                     * Gets or sets whether to display the function's help tip.
                     * @param {boolean} value Whether to display the function's help tip when editing.
                     * @returns {boolean} If no value is set, returns whether the text box displays the function's help tip when editing; otherwise, there is no return value.
                     */
                    showHelp(value?:  boolean): any;
                    /**
                     * Gets or sets the text.
                     * @param {string} value The text.
                     * @returns {string} If no value is set, returns the text; otherwise, there is no return value.
                     */
                    text(value?:  string): any;
                    /**
                     * Removes the binding of an event.
                     * @param {string} type The event type.
                     * @param {Function} fn Specifies the function for which to remove the binding.
                     */
                    unbind(type:  string,  fn:  Function): void;
                    /**
                     * Removes the binding of all events.
                     */
                    unbindAll(): void;
                    /**
                     * Gets or sets the Workbook component to work with the formula text box.
                     * @param {object} value The Workbook component.
                     * @returns {object} If no value is set, returns the workbook component; otherwise, there is no return value.
                     */
                    workbook(value?:  Object): any;
                }
            }

            module Outlines{
                /**
                 * Specifies the status of an outline (range group) summary row or column position.
                 * @enum {number}
                 */
                export enum OutlineDirection{
                    /** The summary row is above or to the left of the group detail.
                     * @type {number}
                     */
                    backward= 0,
                    /** The summary row is below or to the right of the group detail.
                     * @type {number}
                     */
                    forward= 1
                }

                /**
                 * Specifies the status of an outline (range group).
                 * @enum {number}
                 */
                export enum OutlineState{
                    /** Indicates expanded status with the minus sign.
                     * @type {number}
                     */
                    expanded= 0,
                    /** Indicates collapsed status with the plus sign.
                     * @type {number}
                     */
                    collapsed= 1
                }


                export class Outline{
                    /**
                     * Represents an outline (range group) for the worksheet.
                     * @param {number} count The number of rows or columns.
                     * @class
                     */
                    constructor(count:  number);
                    /**
                     * Gets or sets the outline's (range group) direction.
                     * @param {GC.Spread.Sheets.Outlines.OutlineDirection} direction The outline's (range group) direction.
                     * @returns {GC.Spread.Sheets.Outlines.OutlineDirection | GC.Spread.Sheets.Outlines.Outline} If no value is set, returns the outline's (range group) direction; otherwise, returns the outline.
                     */
                    direction(direction?:  GC.Spread.Sheets.Outlines.OutlineDirection): any;
                    /**
                     * Expands all outlines (range groups), using the specified level.
                     * @param {number} level The level of the outline to expand or collapse.
                     * @param {boolean} expand Whether to expand the groups.
                     */
                    expand(level:  number,  expand:  boolean): void;
                    /**
                     * Expands or collapses the specified outline (range group) of rows or columns.
                     * @param {GC.Spread.Sheets.Outlines.OutlineInfo} groupInfo The group information of the range group.
                     * @param {boolean} expand Whether to expand the groups.
                     */
                    expandGroup(groupInfo:  GC.Spread.Sheets.Outlines.OutlineInfo,  expand:  boolean): void;
                    /**
                     * Gets the outline (range group) with the specified group level and row or column index.
                     * @param {number} index The index of the row or column.
                     * @param {number} level The level of the outline (range group).
                     * @returns {GC.Spread.Sheets.Outlines.OutlineInfo} The specified range group.
                     */
                    find(index:  number,  level:  number): GC.Spread.Sheets.Outlines.OutlineInfo;
                    /**
                     * Gets the collapsed internal.
                     * @param {number} index The index.
                     * @returns {boolean} <c>true</c> if collapsed; otherwise, <c>false</c>.
                     */
                    getCollapsed(index:  number): boolean;
                    /**
                     * Gets the level of a specified row or column.
                     * The level's index is zero-based.
                     * @param {number} index The index of the row or column.
                     * @returns {number} The level for the row or column.
                     */
                    getLevel(index:  number): number;
                    /**
                     * Gets the number of the deepest level.
                     * @remarks The level index is zero-based.
                     * @returns {number} The number of the deepest level.
                     */
                    getMaxLevel(): number;
                    /**
                     * Gets the state for the specified group.
                     * @param {GC.Spread.Sheets.Outlines.OutlineInfo} groupInfo The group information.
                     * @returns {GC.Spread.Sheets.Outlines.OutlineState} The group state.
                     */
                    getState(groupInfo:  GC.Spread.Sheets.Outlines.OutlineInfo): GC.Spread.Sheets.Outlines.OutlineState;
                    /**
                     * Groups a range of rows or columns into an outline (range group) from a specified start index.
                     * @param {number} index The group starting index.
                     * @param {number} count The number of rows or columns to group.
                     */
                    group(index:  number,  count:  number): void;
                    /**
                     * Determines whether the range group at the specified index is collapsed.
                     * @param {number} index The index of the row or column in the range group.
                     * @returns {boolean} <c>true</c> if the specified row or column is collapsed; otherwise, <c>false</c>.
                     */
                    isCollapsed(index:  number): boolean;
                    /**
                     * Determines whether the specified index is the end of the group.
                     * @param {number} index The index.
                     * @param {number} processLevel The process level.
                     * @returns {boolean} <c>true</c> if the specfied index is the end of the group; otherwise, <c>false</c>.
                     */
                    isGroupEnd(index:  number,  processLevel:  number): boolean;
                    /**
                     * Refreshes this range group.
                     */
                    refresh(): any;
                    /**
                     * Resumes the adding.
                     */
                    resumeAdding(): any;
                    /**
                     * Sets the collapsed level.
                     * @param {number} index The index.
                     * @param {boolean} collapsed Set to <c>true</c> to collapse the level.
                     */
                    setCollapsed(index:  number,  collapsed:  boolean): void;
                    /**
                     * Suspends the adding.
                     */
                    suspendAdding(): any;
                    /**
                     * Removes all outlines (range groups).
                     */
                    ungroup(): any;
                    /**
                     * Removes a range of rows or columns from the outline (range group) at the specified start index.
                     * @param {number} index The group starting index.
                     * @param {number} count The number of rows or columns to remove.
                     */
                    ungroupRange(index:  number,  count:  number): void;
                }

                export class OutlineInfo{
                    /**
                     * Represents the outline (range group) information.
                     * @param {GC.Spread.Sheets.Outlines.Outline} model The owner of the outline.
                     * @param {number} start The start index of the outline.
                     * @param {number} end The end index of the outline.
                     * @param {number} level The level of the outline.
                     * @class
                     */
                    constructor(model:  GC.Spread.Sheets.Outlines.Outline,  start:  number,  end:  number,  level:  number);
                    /** The children of the group.
                     * @type {Array}
                     */
                    children: any[];
                    /** The end index of the group.
                     * @type {number}
                     */
                    end: number;
                    /** The level of the group.
                     * @type {number}
                     */
                    level: number;
                    /** The owner of the group.
                     * @type {GC.Spread.Sheets.Outlines.Outline}
                     */
                    model: GC.Spread.Sheets.Outlines.Outline;
                    /** The parent of the group.
                     * @type {GC.Spread.Sheets.Outlines.OutlineInfo}
                     */
                    parent: GC.Spread.Sheets.Outlines.OutlineInfo;
                    /** The start index of the group.
                     * @type {number}
                     */
                    start: number;
                    /**
                     * Adds the child.
                     * @param {object} child The child.
                     */
                    addChild(child:  Object): void;
                    /**
                     * Compares this instance to a specified OutlineInfo object and returns an indication of their relative values.
                     * @param {number} index The index of the group item.
                     * @returns {boolean} <c>true</c> if the range group contains the specified index; otherwise, <c>false</c>.
                     */
                    contains(index:  number): boolean;
                    /**
                     * Gets or sets the state of this outline (range group).
                     * @param {GC.Spread.Sheets.Outlines.OutlineState} value The state of this outline (range group).
                     * @returns {GC.Spread.Sheets.Outlines.OutlineState} The state of this outline (range group).
                     */
                    state(value?:  GC.Spread.Sheets.Outlines.OutlineState): GC.Spread.Sheets.Outlines.OutlineState;
                }
            }

            module Print{

                export interface PrintMargins{
                    top: number;
                    bottom: number;
                    left: number;
                    right: number;
                    header: number;
                    footer: number;
                }


                export interface PrintSize{
                    height: number;
                    width: number;
                }

                /**
                 * Specifies the paper kind for the printed page.
                 * @enum {number}
                 */
                export enum PaperKind{
                    /**
                     * Specifies the paper size is 420 mm * 594 mm.
                     */
                    a2= 0x42,
                    
                     /**
                     * Specifies the paper size is 297 mm * 420 mm.
                     */
                    a3= 8,
                    
                     /**
                     * Specifies the paper size is 322 mm * 445 mm.
                     */
                    a3Extra= 0x3f,
                    
                     /**
                     * Specifies the paper size is 322 mm * 445 mm.
                     */
                    a3ExtraTransverse= 0x44,
                    
                     /**
                     * Specifies the paper size is 420 mm * 297 mm.
                     */
                    a3Rotated= 0x4c,
                    
                     /**
                     * Specifies the paper size is 297 mm * 420 mm.
                     */
                    a3Transverse= 0x43,
                    
                     /**
                     * Specifies the paper size is 210 mm * 297 mm.
                     */
                    a4= 9,
                    
                     /**
                     * Specifies the paper size is 236 mm * 322 mm.
                     */
                    a4Extra= 0x35,
                    
                     /**
                     * Specifies the paper size is 210 mm * 330 mm.
                     */
                    a4Plus= 60,
                    
                     /**
                     * Specifies the paper size is 297 mm * 210 mm.
                     */
                    a4Rotated= 0x4d,
                    
                     /**
                     * Specifies the paper size is 210 mm * 297 mm.
                     */
                    a4Small= 10,
                    
                     /**
                     * Specifies the paper size is 210 mm * 297 mm.
                     */
                    a4Transverse= 0x37,
                    
                     /**
                     * Specifies the paper size is 148 mm * 210 mm.
                     */
                    a5= 11,
                    
                     /**
                     * Specifies the paper size is 174 mm * 235 mm.
                     */
                    a5Extra= 0x40,
                    
                     /**
                     * Specifies the paper size is 210 mm * 148 mm.
                     */
                    a5Rotated= 0x4e,
                    
                     /**
                     * Specifies the paper size is 148 mm * 210 mm.
                     */
                    a5Transverse= 0x3d,
                    
                     /**
                     * Specifies the paper size is 105 mm * 148 mm.
                     */
                    a6= 70,
                    
                     /**
                     * Specifies the paper size is 148 mm * 105 mm.
                     */
                    a6Rotated= 0x53,
                    
                     /**
                     * Specifies the paper size is 227 mm * 356 mm.
                     */
                    aPlus= 0x39,
                    
                     /**
                     * Specifies the paper size is 250 mm * 353 mm.
                     */
                    b4= 12,
                    
                     /**
                     * Specifies the paper size is 250 mm * 353 mm.
                     */
                    b4Envelope= 0x21,
                    
                     /**
                     * Specifies the paper size is 364 mm * 257 mm.
                     */
                    b4JisRotated= 0x4f,
                    
                     /**
                     * Specifies the paper size is 176 mm * 250 mm.
                     */
                    b5= 13,
                    
                     /**
                     * Specifies the paper size is 176 mm * 250 mm.
                     */
                    b5Envelope= 0x22,
                    
                     /**
                     * Specifies the paper size is 201 mm * 276 mm.
                     */
                    b5Extra= 0x41,
                    
                     /**
                     * Specifies the paper size is 257 mm * 182 mm.
                     */
                    b5JisRotated= 80,
                    
                     /**
                     * Specifies the paper size is 182 mm * 257 mm.
                     */
                    b5Transverse= 0x3e,
                    
                     /**
                     * Specifies the paper size is 176 mm * 125 mm.
                     */
                    b6Envelope= 0x23,
                    
                     /**
                     * Specifies the paper size is 128 mm * 182 mm.
                     */
                    b6Jis= 0x58,
                    
                     /**
                     * Specifies the paper size is 182 mm * 128 mm.
                     */
                    b6JisRotated= 0x59,
                    
                     /**
                     * Specifies the paper size is 305 mm * 487 mm.
                     */
                    bPlus= 0x3a,
                    
                     /**
                     * Specifies the paper size is 324 mm * 458 mm.
                     */
                    c3Envelope= 0x1d,
                    
                     /**
                     * Specifies the paper size is 229 mm * 324 mm.
                     */
                    c4Envelope= 30,
                    
                     /**
                     * Specifies the paper size is 162 mm * 229 mm.
                     */
                    c5Envelope= 0x1c,
                    
                     /**
                     * Specifies the paper size is 114 mm * 229 mm.
                     */
                    c65Envelope= 0x20,
                    
                     /**
                     * Specifies the paper size is 114 mm * 162 mm.
                     */
                    c6Envelope= 0x1f,
                    
                     /**
                     * Specifies the paper size is 17 in. * 22 in.
                     */
                    cSheet= 0x18,
                    
                     /**
                     * Specifies the paper size is defined by the user.
                     */
                    custom= 0,
                    
                     /**
                     * Specifies the paper size is 110 mm * 220 mm.
                     */
                    dlEnvelope= 0x1b,
                    
                     /**
                     * Specifies the paper size is 22 in. * 34 in.
                     */
                    dSheet= 0x19,
                    
                     /**
                     * Specifies the paper size is 34 in. * 44 in.
                     */
                    eSheet= 0x1a,
                    
                     /**
                     * Specifies the paper size is 7.25 in. * 10.5 in.
                     */
                    executive= 7,
                    
                     /**
                     * Specifies the paper size is 8.5 in. * 13 in.
                     */
                    folio= 14,
                    
                     /**
                     * Specifies the paper size is 8.5 in. * 13 in.
                     */
                    germanLegalFanfold= 0x29,
                    
                     /**
                     * Specifies the paper size is 8.5 in. * 12 in.
                     */
                    germanStandardFanfold= 40,
                    
                     /**
                     * Specifies the paper size is 220 mm * 220 mm.
                     */
                    inviteEnvelope= 0x2f,
                    
                     /**
                     * Specifies the paper size is 250 mm * 353 mm.
                     */
                    isoB4= 0x2a,
                    
                     /**
                     * Specifies the paper size is 110 mm * 230 mm.
                     */
                    italyEnvelope= 0x24,
                    
                     /**
                     * Specifies the paper size is 200 mm * 148 mm.
                     */
                    japaneseDoublePostcard= 0x45,
                    
                     /**
                     * Specifies the paper size is 148 mm * 200 mm.
                     */
                    japaneseDoublePostcardRotated= 0x52,
                    
                     /**
                     * Specifies the paper size is Japanese Chou #3 envelope, 120 mm * 235 mm.
                     */
                    japaneseEnvelopeChouNumber3= 0x49,
                    
                     /**
                     * Specifies the paper size is Japanese rotated Chou #3 envelope, 235 mm * 120 mm.
                     */
                    japaneseEnvelopeChouNumber3Rotated= 0x56,
                    
                     /**
                     * Specifies the paper size is Japanese Chou #4 envelope, 90 mm * 205 mm.
                     */
                    japaneseEnvelopeChouNumber4= 0x4a,
                    
                     /**
                     * Specifies the paper size is Japanese rotated Chou #4 envelope, 205 mm * 90 mm.
                     */
                    japaneseEnvelopeChouNumber4Rotated= 0x57,
                    
                     /**
                     * Specifies the paper size is Japanese Kaku #2 envelope, 240 mm * 332 mm.
                     */
                    japaneseEnvelopeKakuNumber2= 0x47,
                    
                     /**
                     * Specifies the paper size is Japanese rotated Kaku #2 envelope, 332 mm * 240 mm.
                     */
                    japaneseEnvelopeKakuNumber2Rotated= 0x54,
                    
                     /**
                     * Specifies the paper size is Japanese Kaku #3 envelope, 216 mm * 277 mm.
                     */
                    japaneseEnvelopeKakuNumber3= 0x48,
                    
                     /**
                     * Specifies the paper size is Japanese rotated Kaku #3 envelope, 277 mm * 216 mm.
                     */
                    japaneseEnvelopeKakuNumber3Rotated= 0x55,
                    
                     /**
                     * Specifies the paper size is Japanese You #4 envelope, 235 mm * 105 mm.
                     */
                    japaneseEnvelopeYouNumber4= 0x5b,
                    
                     /**
                     * Specifies the paper size is Japanese You #4 rotated envelope, 105 mm * 235 mm.
                     */
                    japaneseEnvelopeYouNumber4Rotated= 0x5c,
                    
                     /**
                     * Specifies the paper size is 100 mm * 148 mm.
                     */
                    japanesePostcard= 0x2b,
                    
                     /**
                     * Specifies the paper size is 148 mm * 100 mm.
                     */
                    japanesePostcardRotated= 0x51,
                    
                     /**
                     * Specifies the paper size is 17 in. * 11 in.
                     */
                    ledger= 4,
                    
                     /**
                     * Specifies the paper size is 8.5 in. * 14 in.
                     */
                    legal= 5,
                    
                     /**
                     * Specifies the paper size is legal extra paper (9.275 in. * 15 in.).
                     * This value is specific to the PostScript driver and is used only by Linotronic printers in order to conserve paper.
                     */
                    legalExtra= 0x33,
                    
                     /**
                     * Specifies the paper size is 8.5 in. * 11 in.
                     */
                    letter= 1,
                    
                     /**
                     * Specifies the paper size is letter extra paper (9.275 in. * 12 in.).
                     * This value is specific to the PostScript driver and is used only by Linotronic printers in order to conserve paper.
                     */
                    letterExtra= 50,
                    
                     /**
                     * Specifies the paper size 9.275 in. * 12 in.
                     */
                    letterExtraTransverse= 0x38,
                    
                     /**
                     * Specifies the paper size is 8.5 in. * 12.69 in.
                     */
                    letterPlus= 0x3b,
                    
                     /**
                     * Specifies the paper size is 11 in. * 8.5 in.
                     */
                    letterRotated= 0x4b,
                    
                     /**
                     * Specifies the paper size is 8.5 in. * 11 in.
                     */
                    letterSmall= 2,
                    
                     /**
                     * Specifies the paper size is 8.275 in. * 11 in.
                     */
                    letterTransverse= 0x36,
                    
                     /**
                     * Specifies the paper size is 3.875 in. * 7.5 in.
                     */
                    monarchEnvelope= 0x25,
                    
                     /**
                     * Specifies the paper size is 8.5 in. * 11 in.
                     */
                    note= 0x12,
                    
                     /**
                     * Specifies the paper size is 4.125 in. * 9.5 in.
                     */
                    number10Envelope= 20,
                    
                     /**
                     * Specifies the paper size is 4.5 in. * 10.375 in.
                     */
                    number11Envelope= 0x15,
                    
                     /**
                     * Specifies the paper size is 4.75 in. * 11 in.
                     */
                    number12Envelope= 0x16,
                    
                     /**
                     * Specifies the paper size is 5 in. * 11.5 in.
                     */
                    number14Envelope= 0x17,
                    
                     /**
                     * Specifies the paper size is 3.875 in. * 8.875 in.
                     */
                    number9Envelope= 0x13,
                    
                     /**
                     * Specifies the paper size is 3.625 in. * 6.5 in.
                     */
                    personalEnvelope= 0x26,
                    
                     /**
                     * Specifies the paper size is 146 mm * 215 mm.
                     */
                    prc16K= 0x5d,
                    
                     /**
                     * Specifies the paper size is 146 mm * 215 mm.
                     */
                    prc16KRotated= 0x6a,
                    
                     /**
                     * Specifies the paper size is 97 mm * 151 mm.
                     */
                    prc32K= 0x5e,
                    
                     /**
                     * Specifies the paper size is 97 mm * 151 mm.
                     */
                    prc32KBig= 0x5f,
                    
                     /**
                     * Specifies the paper size is 97 mm * 151 mm.
                     */
                    prc32KBigRotated= 0x6c,
                    
                     /**
                     * Specifies the paper size is 97 mm * 151 mm.
                     */
                    prc32KRotated= 0x6b,
                    
                     /**
                     * Specifies the paper size is 102 mm * 165 mm.
                     */
                    prcEnvelopeNumber1= 0x60,
                    
                     /**
                     * Specifies the paper size is 324 mm * 458 mm.
                     */
                    prcEnvelopeNumber10= 0x69,
                    
                     /**
                     * Specifies the paper size is 458 mm * 324 mm.
                     */
                    prcEnvelopeNumber10Rotated= 0x76,
                    
                     /**
                     * Specifies the paper size is 165 mm * 102 mm.
                     */
                    prcEnvelopeNumber1Rotated= 0x6d,
                    
                     /**
                     * Specifies the paper size is 102 mm * 176 mm.
                     */
                    prcEnvelopeNumber2= 0x61,
                    
                     /**
                     * Specifies the paper size is 176 mm * 102 mm.
                     */
                    prcEnvelopeNumber2Rotated= 110,
                    
                     /**
                     * Specifies the paper size is 125 mm * 176 mm.
                     */
                    prcEnvelopeNumber3= 0x62,
                    
                     /**
                     * Specifies the paper size is 176 mm * 125 mm.
                     */
                    prcEnvelopeNumber3Rotated= 0x6f,
                    
                     /**
                     * Specifies the paper size is 110 mm * 208 mm.
                     */
                    prcEnvelopeNumber4= 0x63,
                    
                     /**
                     * Specifies the paper size is 208 mm * 110 mm.
                     */
                    prcEnvelopeNumber4Rotated= 0x70,
                    
                     /**
                     * Specifies the paper size is 110 mm * 220 mm.
                     */
                    prcEnvelopeNumber5= 100,
                    
                     /**
                     * Specifies the paper size is 220 mm * 110 mm.
                     */
                    prcEnvelopeNumber5Rotated= 0x71,
                    
                     /**
                     * Specifies the paper size is 120 mm * 230 mm.
                     */
                    prcEnvelopeNumber6= 0x65,
                    
                     /**
                     * Specifies the paper size is 230 mm * 120 mm.
                     */
                    prcEnvelopeNumber6Rotated= 0x72,
                    
                     /**
                     * Specifies the paper size is 160 mm * 230 mm.
                     */
                    prcEnvelopeNumber7= 0x66,
                    
                     /**
                     * Specifies the paper size is 230 mm * 160 mm.
                     */
                    prcEnvelopeNumber7Rotated= 0x73,
                    
                     /**
                     * Specifies the paper size is 120 mm * 309 mm.
                     */
                    prcEnvelopeNumber8= 0x67,
                    
                     /**
                     * Specifies the paper size is 309 mm * 120 mm.
                     */
                    prcEnvelopeNumber8Rotated= 0x74,
                    
                     /**
                     * Specifies the paper size is 229 mm * 324 mm.
                     */
                    prcEnvelopeNumber9= 0x68,
                    
                     /**
                     * Specifies the paper size is 324 mm * 229 mm.
                     */
                    prcEnvelopeNumber9Rotated= 0x75,
                    
                     /**
                     * Specifies the paper size is 215 mm * 275 mm.
                     */
                    quarto= 15,
                    
                     /**
                     * Specifies the paper size is 10 in. * 11 in.
                     */
                    standard10x11= 0x2d,
                    
                     /**
                     * Specifies the paper size is 10 in. * 14 in.
                     */
                    standard10x14= 0x10,
                    
                     /**
                     * Specifies the paper size is 11 in. * 17 in.
                     */
                    standard11x17= 0x11,
                    
                     /**
                     * Specifies the paper size is 12 in. * 11 in.
                     */
                    standard12x11= 90,
                    
                     /**
                     * Specifies the paper size is 15 in. * 11 in.
                     */
                    standard15x11= 0x2e,
                    
                     /**
                     * Specifies the paper size is 9 in. * 11 in.
                     */
                    standard9x11= 0x2c,
                    
                     /**
                     * Specifies the paper size is 5.5 in. * 8.5 in.
                     */
                    statement= 6,
                    
                     /**
                     * Specifies the paper size is 11 in. * 17 in.
                     */
                    tabloid= 3,
                    
                     /**
                     * Specifies the paper size is 11.69 in. * 18 in.
                     */
                    tabloidExtra= 0x34,
                    
                     /**
                     * Specifies the paper size is 14.875 in. * 11 in.
                     */
                    usStandardFanfold= 0x27
                }

                /**
                 * Specifics the type of centering for the printed page.
                 * @enum {number}
                 */
                export enum PrintCentering{
                    /**
                     * Does not center the printed page at all.
                     */
                    none= 0,
                    /**
                     * Centers the printed layout horizontally on the page.
                     */
                    horizontal= 1,
                    /**
                     * Centers the printed layout vertically on the page.
                     */
                    vertical= 2,
                    /**
                     * Centers the printed layout both horizontally and vertically on the page.
                     */
                    both= 3
                }

                /**
                 * Specifies the order in which pages are printed.
                 * @enum {number}
                 */
                export enum PrintPageOrder{
                    /**
                     * Automatically determines the best order for printing pages.
                     */
                    auto= 0,
                    /**
                     * Prints pages down then across.
                     */
                    downThenOver= 1,
                    /**
                     * Prints pages across then down.
                     */
                    overThenDown= 2
                }

                /**
                 * Specifies the page orientation used for printing.
                 * @enum {number}
                 */
                export enum PrintPageOrientation{
                    /**
                     * Prints portrait orientation.
                     */
                    portrait= 1,
                    /**
                     * Prints landscape orientation.
                     */
                    landscape= 2
                }

                /**
                 * Specifies whether the area is visible.
                 * @enum {number}
                 */
                export enum PrintVisibilityType{
                    /**
                     * Inherits the setting from the Worksheet class.
                     */
                    inherit= 0,
                    /**
                     * Hides the area.
                     */
                    hide= 1,
                    /**
                     * Shows in each page.
                     */
                    show= 2,
                    /**
                     * Shows once.
                     */
                    showOnce= 3
                }


                export class PaperSize{
                    /**
                     * Specifies the paper size.<br />
                     * The constructor has 3 modes.<br />
                     * If there are 2 parameters, the parameters are width and height with a type of number;<br />
                     * If there is 1 parameter, the parameter is kind which is a GC.Spread.Sheets.Print.PaperKind type;<br />
                     * If there is no parameter, the result is the same as the second mode and the kind option is GC.Spread.Sheets.Print.PaperKind.letter.
                     * @class
                     * @param {number|GC.Spread.Sheets.Print.PaperKind} widthOrKind The width of the paper, in hundredths of an inch; or the kind of the paper and the type is GC.Spread.Sheets.Print.PaperKind.
                     * @param {number} height The height of the paper, in hundredths of an inch.
                     */
                    constructor(widthOrKind?:  any,  height?:  number);
                    /**
                     * Gets the paper size, in hundredths of an inch.
                     * @param {GC.Spread.Sheets.Print.PaperKind} kind The kind of the paper.
                     * @returns {Object} The size which contains width and height of the paper.
                     * size.width {number} The width of the size, in hundredths of an inch.
                     * size.height {number} The height of the size, in hundredths of an inch.
                     */
                    getPageSize(kind:  GC.Spread.Sheets.Print.PaperKind): GC.Spread.Sheets.Print.PrintSize;
                    /**
                     * Gets or sets the height of the paper, in hundredths of an inch.
                     * @param {number} value The height of the paper.
                     * @returns {number | GC.Spread.Sheets.Print.PaperSize} If no value is set, returns the height of the paper; otherwise, returns the paper size.
                     */
                    height(value?:  number): any;
                    /**
                     * Gets or sets the kind of the paper.
                     * @param {GC.Spread.Sheets.Print.PaperKind} value The kind of the paper.
                     * @returns {GC.Spread.Sheets.Print.PaperKind | GC.Spread.Sheets.Print.PaperSize} If no value is set, returns the kind of the paper; otherwise, returns the paper size.
                     */
                    kind(value?:  GC.Spread.Sheets.Print.PaperKind): any;
                    /**
                     * Gets or sets the width of the paper, in hundredths of an inch.
                     * @param {number} value The width of the paper.
                     * @returns {number | GC.Spread.Sheets.Print.PaperSize} If no value is set, returns the width of the paper; otherwise, returns the paper size.
                     */
                    width(value?:  number): any;
                }

                export class PrintInfo{
                    /**
                     * Represents the information to use when printing a Worksheet.
                     * @class
                     */
                    constructor();
                    /**
                     * Gets or sets whether column widths are adjusted to fit the longest text width for printing.
                     * @param {boolean} value Whether column widths are adjusted to fit the longest text width for printing.
                     * @returns {boolean | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns whether column widths are adjusted to fit the longest text width for printing; otherwise, returns the print setting information.
                     */
                    bestFitColumns(value?:  boolean): any;
                    /**
                     * Gets or sets whether row heights are adjusted to fit the tallest text height for printing.
                     * @param {boolean} value Whether row heights are adjusted to fit the tallest text height for printing.
                     * @returns {boolean | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns whether row heights are adjusted to fit the tallest text height for printing; otherwise, returns the print setting information.
                     */
                    bestFitRows(value?:  boolean): any;
                    /**
                     * Gets or sets whether to print in black and white.
                     * @param {boolean} value Whether to print in black and white.
                     * @returns {boolean | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns whether to print in black and white; otherwise, returns the print setting information.
                     */
                    blackAndWhite(value?:  boolean): any;
                    /**
                     * Gets or sets how the printed page is centered.
                     * @param {GC.Spread.Sheets.Print.PrintCentering} value How the printed page is centered.
                     * @returns {GC.Spread.Sheets.Print.PrintCentering | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns how the printed page is centered; otherwise, returns the print setting information.
                     */
                    centering(value?:  GC.Spread.Sheets.Print.PrintCentering): any;
                    /**
                     * Gets or sets the last column to print when printing a cell range.
                     * @param {number} value The last column to print when printing a cell range.
                     * @returns {number | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the last column to print when printing a cell range; otherwise, returns the print setting information.
                     */
                    columnEnd(value?:  number): any;
                    /**
                     * Gets or sets the first column to print when printing a cell range.
                     * @param {number} value The first column to print when printing a cell range.
                     * @returns {number | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the first column to print when printing a cell range; otherwise, returns the print setting information.
                     */
                    columnStart(value?:  number): any;
                    /**
                     * Gets or sets the page number to print on the first page.
                     * @param {number} value The page number to print on the first page.
                     * @returns {number | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the page number to print on the first page; otherwise, returns the print setting information.
                     */
                    firstPageNumber(value?:  number): any;
                    /**
                     * Gets or sets the number of vertical pages to check when optimizing printing.
                     * @param {number} value The number of vertical pages to check when optimizing printing.
                     * @returns {number | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the number of vertical pages to check; otherwise, returns the print setting information.
                     */
                    fitPagesTall(value?:  number): any;
                    /**
                     * Gets or sets the number of horizontal pages to check when optimizing the printing.
                     * @param {number} value The number of horizontal pages to check when optimizing the printing.
                     * @returns {number | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the number of horizontal pages to check; otherwise, returns the print setting information.
                     */
                    fitPagesWide(value?:  number): any;
                    /**
                     * Gets or sets the text and format of the center footer on printed pages.
                     * @param {string} value The text and format of the center footer on printed pages.
                     * @returns {string | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the text and format of the center footer on printed pages; otherwise, returns the print setting information.
                     */
                    footerCenter(value?:  string): any;
                    /**
                     * Gets or sets the image for the center section of the footer.
                     * @param {string} value The image for the center section of the footer.
                     * @returns {string | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the image for the center section of the footer; otherwise, returns the print setting information.
                     */
                    footerCenterImage(value?:  string): any;
                    /**
                     * Gets or sets the text and format of the left footer on printed pages.
                     * @param {string} value The text and format of the left footer on printed pages.
                     * @returns {string | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the text and format of the left footer on printed pages; otherwise, returns the print setting information.
                     */
                    footerLeft(value?:  string): any;
                    /**
                     * Gets or sets the image for the left section of the footer.
                     * @param {string} value The image for the left section of the footer.
                     * @returns {string | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the image for the left section of the footer; otherwise, returns the print setting information.
                     */
                    footerLeftImage(value?:  string): any;
                    /**
                     * Gets or sets the text and format of the right footer on printed pages.
                     * @param {string} value The text and format of the right footer on printed pages.
                     * @returns {string | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the text and format of the right footer on printed pages; otherwise, returns the print setting information.
                     */
                    footerRight(value?:  string): any;
                    /**
                     * Gets or sets the image for the right section of the footer.
                     * @param {string} value The image for the right section of the footer.
                     * @returns {string | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the image for the right section of the footer; otherwise, returns the print setting information.
                     */
                    footerRightImage(value?:  string): any;
                    /**
                     * Gets or sets the text and format of the center header on printed pages.
                     * @param {string} value The text and format of the center header on printed pages.
                     * @returns {string | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the text and format of the center header on printed pages; otherwise, returns the print setting information.
                     */
                    headerCenter(value?:  string): any;
                    /**
                     * Gets or sets the image for the center section of the header.
                     * @param {string} value The image for the center section of the header.
                     * @returns {string | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the image for the center section of the header; otherwise, returns the print setting information.
                     */
                    headerCenterImage(value?:  string): any;
                    /**
                     * Gets or sets the text and format of the left header on printed pages.
                     * @param {string} value The text and format of the left header on printed pages.
                     * @returns {string | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the text and format of the left header on printed pages; otherwise, returns the print setting information.
                     */
                    headerLeft(value?:  string): any;
                    /**
                     * Gets or sets the image for the left section of the header.
                     * @param {string} value The image for the left section of the header.
                     * @returns {string | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the image for the left section of the header; otherwise, returns the print setting information.
                     */
                    headerLeftImage(value?:  string): any;
                    /**
                     * Gets or sets the text and format of the right header on printed pages.
                     * @param {string} value The text and format of the right header on printed pages.
                     * @returns {string | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the text and format of the right header on printed pages; otherwise, returns the print setting information.
                     */
                    headerRight(value?:  string): any;
                    /**
                     * Gets or sets the image for the right section of the header.
                     * @param {string} value The image for the right section of the header.
                     * @returns {string | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the image for the right section of the header; otherwise, returns the print setting information.
                     */
                    headerRightImage(value?:  string): any;
                    /**
                     * Gets or sets the margins for printing, in hundredths of an inch.
                     * @param {Object} value The margins for printing.<br />
                     * value.top {number} top The top margin, in hundredths of an inch.<br />
                     * value.bottom {number} bottom The bottom margin, in hundredths of an inch.<br />
                     * value.left {number} left The left margin, in hundredths of an inch.<br />
                     * value.right {number} right The right margin, in hundredths of an inch.<br />
                     * value.header {number} header The header offset, in hundredths of an inch.<br />
                     * value.footer {number} footer The footer offset, in hundredths of an inch.
                     * @returns {Object | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the margins for printing; otherwise, returns the print setting information.
                     */
                    margin(value?:  GC.Spread.Sheets.Print.PrintMargins): any;
                    /**
                     * Gets or sets the page orientation used for printing.
                     * @param {GC.Spread.Sheets.Print.PrintPageOrientation} value The page orientation used for printing.
                     * @returns {GC.Spread.Sheets.Print.PrintPageOrientation | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the page orientation used for printing; otherwise, returns the print setting information.
                     */
                    orientation(value?:  GC.Spread.Sheets.Print.PrintPageOrientation): any;
                    /**
                     * Gets or sets the order in which pages print.
                     * @param {GC.Spread.Sheets.Print.PrintPageOrder} value The order in which pages print.
                     * @returns {GC.Spread.Sheets.Print.PrintPageOrder | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns a value that specifies the order in which pages print; otherwise, returns the print setting information.
                     */
                    pageOrder(value?:  GC.Spread.Sheets.Print.PrintPageOrder): any;
                    /**
                     * Gets or sets the page range for printing.
                     * @param {string} value The page numbers or page ranges separated by commas counting from the beginning of the document. For example, type "1,3,5-12".
                     * @returns {string | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns a string that provides page numbers or page ranges; otherwise, returns the print setting information.}
                     */
                    pageRange(value?:  string): any;
                    /**
                     * Gets or sets the paper size for printing.
                     * @param {GC.Spread.Sheets.Print.PaperSize} value The paper size for printing.<br />
                     * value.width {number} The width, in hundredths of an inch.<br />
                     * value.height {number} The height, in hundredths of an inch.
                     * @returns {GC.Spread.Sheets.Print.PaperSize | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the paper size for printing; otherwise, returns the print setting information.
                     */
                    paperSize(value?:  GC.Spread.Sheets.Print.PaperSize): any;
                    /**
                     * Gets or sets the quality factor for printing.
                     * @param {number} value The quality factor for printing is a positive integer between 1 and 8. The greater the quality factor, the better the printing quality. When the quality factor is bigger, the printing efficiency is affected.
                     * @returns {number | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the quality factor for printing; otherwise, returns the print setting information.
                     */
                    qualityFactor(value?:  number): any;
                    /**
                     * Gets or sets the last column of a range of columns to print on the left of each page.
                     * @param {number} value The last column of a range of columns to print on the left of each page.
                     * @returns {number | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the last column of a range of columns to print on the left of each page; otherwise, returns the print setting information.
                     */
                    repeatColumnEnd(value?:  number): any;
                    /**
                     * Gets or sets the first column of a range of columns to print on the left of each page.
                     * @param {number} value The first column of a range of columns to print on the left of each page.
                     * @returns {number | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the first column of a range of columns to print on the left of each page; otherwise, returns the print setting information.
                     */
                    repeatColumnStart(value?:  number): any;
                    /**
                     * Gets or sets the last row of a range of rows to print at the top of each page.
                     * @param {number} value The last row of a range of rows to print at the top of each page.
                     * @returns {number | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the last row of a range of rows to print at the top of each page; otherwise, returns the print setting information.
                     */
                    repeatRowEnd(value?:  number): any;
                    /**
                     * Gets or sets the first row of a range of rows to print at the top of each page.
                     * @param {number} value The first row of a range of rows to print at the top of each page.
                     * @returns {number | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the first row of a range of rows to print at the top of each page; otherwise, returns the print setting information.
                     */
                    repeatRowStart(value?:  number): any;
                    /**
                     * Gets or sets the last row to print when printing a cell range.
                     * @param {number} value The last row to print when printing a cell range.
                     * @returns {number | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the last row to print when printing a cell range; otherwise, returns the print setting information.
                     */
                    rowEnd(value?:  number): any;
                    /**
                     * Gets or sets the first row to print when printing a cell range.
                     * @param {number} value The first row to print when printing a cell range.
                     * @returns {number | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns the first row to print when printing a cell range; otherwise, returns the print setting information.
                     */
                    rowStart(value?:  number): any;
                    /**
                     * Gets or sets whether to print an outline border around the entire control.
                     * @param {boolean} value Whether to print an outline border around the entire control.
                     * @returns {boolean | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns whether to print an outline border around the entire control; otherwise, returns the print setting information.
                     */
                    showBorder(value?:  boolean): any;
                    /**
                     * Gets or sets whether to print the column header.
                     * @param {GC.Spread.Sheets.Print.PrintVisibilityType} value Whether to print the column header.
                     * @returns {GC.Spread.Sheets.Print.PrintVisibilityType | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns whether to print the column header; otherwise, returns the print setting information.
                     */
                    showColumnHeader(value?:  GC.Spread.Sheets.Print.PrintVisibilityType): any;
                    /**
                     * Gets or sets whether to print the grid lines.
                     * @param {boolean} value Whether to print the grid lines.
                     * @returns {boolean | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns whether to print the grid lines; otherwise, returns the print setting information.
                     */
                    showGridLine(value?:  boolean): any;
                    /**
                     * Gets or sets whether to print the row header.
                     * @param {GC.Spread.Sheets.Print.PrintVisibilityType} value Whether to print the row header.
                     * @returns {GC.Spread.Sheets.Print.PrintVisibilityType | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns whether to print the row header; otherwise, returns the print setting information.
                     */
                    showRowHeader(value?:  GC.Spread.Sheets.Print.PrintVisibilityType): any;
                    /**
                     * Gets or sets whether to print only rows and columns that contain data.
                     * @param {boolean} value Whether to print only rows and columns that contain data.
                     * @returns {boolean | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns whether to print only rows and columns that contain data; otherwise, returns the print setting information.
                     */
                    useMax(value?:  boolean): any;
                    /**
                     * Gets or sets the zoom factor used for printing.
                     * @param {number} value The zoom factor used for printing.
                     * @returns {number | GC.Spread.Sheets.Print.PrintInfo} If no value is set, returns a value that specifies the amount to enlarge or reduce the printed worksheet; otherwise, returns the print setting information.
                     */
                    zoomFactor(value?:  number): any;
                }
            }

            module Search{
                /**
                 * Specifies the type of search flags.
                 * @enum {number}
                 */
                export enum SearchFlags{
                    /** none
                     * @type {number}
                     */
                    none= 0,
                    /** Determines whether the search considers the case of the letters in the search string.
                     * @type {number}
                     */
                    ignoreCase= 1,
                    /** Determines whether the search considers only an exact match.
                     * @type {number}
                     */
                    exactMatch= 2,
                    /** Determines whether the search considers wildcard characters (*, ?) in the search string.
                     * @type {number}
                     */
                    useWildCards= 4,
                    /** Determines whether to search within a cell range.
                     * @type {number}
                     */
                    blockRange= 8
                }

                /**
                 * Specifies where the search string is found.
                 * @enum {number}
                 */
                export enum SearchFoundFlags{
                    /**
                     * Indicates that no string is found.
                     * @type {number}
                     */
                    none= 0,
                    /**
                     * Indicates that the string is found in the cell text.
                     * @type {number}
                     */
                    cellText= 1,
                    /**
                     * Indicates that the string is found in the cell tag.
                     * @type {number}
                     */
                    cellTag= 4,
                    /**
                     * Indicates that the string is found in the cell formula.
                     * @type {number}
                     */
                    cellFormula= 8
                }

                /**
                 * Specifies the type of search direction.
                 * @enum {number}
                 */
                export enum SearchOrder{
                    /** Determines whether the search goes by column, row coordinates.
                     * @type {number}
                     */
                    zOrder= 0,
                    /** Determines whether the search goes by row, column coordinates.
                     * @type {number}
                     */
                    nOrder= 1
                }


                export class SearchCondition{
                    /**
                     * Defines the search condition.
                     * @class
                     */
                    constructor();
                    /** The index of the column at which to end.
                     * @type {number}
                     */
                    columnEnd: number;
                    /** The index of the column at which to start.
                     * @type {number}
                     */
                    columnStart: number;
                    /** Index of the sheet on which to end searching.
                     * @type {number}
                     */
                    endSheetIndex: number;
                    /** The index of the row at which to end.
                     * @type {number}
                     */
                    rowEnd: number;
                    /** The index of the row at which to start.
                     * @type {number}
                     */
                    rowStart: number;
                    /** The enumeration that specifies the options of the search.
                     * @type {GC.Spread.Sheets.Search.SearchFlags}
                     */
                    searchFlags: GC.Spread.Sheets.Search.SearchFlags;
                    /** The enumeration that specifies whether the search goes by coordinates of (column, row) or (row, column).
                     * @type {GC.Spread.Sheets.Search.SearchOrder}
                     */
                    searchOrder: GC.Spread.Sheets.Search.SearchOrder;
                    /** The string for which to search.
                     * @type {string}
                     */
                    searchString: string;
                    /** The enumeration that indicates whether the search includes the content in the cell notes, tags, or text.
                     * @type {GC.Spread.Sheets.Search.SearchFoundFlags}
                     */
                    searchTarget: GC.Spread.Sheets.Search.SearchFoundFlags;
                    /** The area of the sheet for search.
                     * @type {GC.Spread.Sheets.SheetArea}
                     */
                    sheetArea: GC.Spread.Sheets.SheetArea;
                    /** Index of the sheet on which to start searching.
                     * @type {number}
                     */
                    startSheetIndex: number;
                }

                export class SearchResult{
                    /**
                     * Defines the search result.
                     * @class
                     */
                    constructor();
                    /** The index of the column at which a match is found.
                     * @type {number}
                     */
                    foundColumnIndex: number;
                    /** The index of the row at which a match is found.
                     * @type {number}
                     */
                    foundRowIndex: number;
                    /** The index of the sheet in which a match is found.
                     * @type {number}
                     */
                    foundSheetIndex: number;
                    /** The found string.
                     * @type {object}
                     */
                    foundString: Object;
                    /**
                     * An enumeration that specifies what is matched.
                     * @type {GC.Spread.Sheets.Search.SearchFoundFlags}
                     */
                    searchFoundFlag: GC.Spread.Sheets.Search.SearchFoundFlags;
                }
            }

            module Slicers{

                export class ItemSlicer{
                    /**
                     * Represents an item slicer.
                     * @param {string} name The name of the item slicer.
                     * @param {GC.Spread.Slicers.GeneralSlicerData} slicerData An instance of the GeneralSlicerData or TableSlicerData.
                     * @param {string} columnName The column name that relates to the item slicer.
                     * @class GC.Spread.Sheets.Slicers.ItemSlicer
                     */
                    constructor(name:  string,  slicerData:  GC.Spread.Slicers.GeneralSlicerData,  columnName:  string);
                    /**
                     * Gets or sets the caption name of the item slicer.
                     * @param {string} value The caption name of the item slicer.
                     * @returns {string | GC.Spread.Sheets.Slicers.ItemSlicer} If no value is set, returns the caption name of the item slicer; otherwise, returns the item slicer.
                     */
                    captionName(value?:  string): any;
                    /**
                     * Gets or sets the column count of the item slicer.
                     * @param {number} value The column count of the item slicer.
                     * @returns {number | GC.Spread.Sheets.Slicers.ItemSlicer} If no value is set, returns the column count of the item slicer; otherwise, returns the item slicer.
                     */
                    columnCount(value?:  number): any;
                    /**
                     * Gets the dom element of the item slicer.
                     * @returns {HTMLElement} The dom element of the item slicer.
                     */
                    getDOMElement(): HTMLElement;
                    /**
                     * Gets or sets the height of the item slicer.
                     * @param {number} value The height of the item slicer.
                     * @returns {number | GC.Spread.Sheets.Slicers.ItemSlicer} If no value is set, returns the height of the item slicer; otherwise, returns the item slicer.
                     */
                    height(value?:  number): any;
                    /**
                     * Gets or sets the item height of the item slicer.
                     * @param {number} value The item height of the item slicer.
                     * @returns {number | GC.Spread.Sheets.Slicers.ItemSlicer} If no value is set, returns the item height of the item slicer; otherwise, returns the item slicer.
                     */
                    itemHeight(value?:  number): any;
                    /**
                     * Gets or sets the name of the item slicer.
                     * @param {string} value The name of the item slicer.
                     * @returns {string | GC.Spread.Sheets.Slicers.ItemSlicer} If no value is set, returns the name of the item slicer; otherwise, returns the item slicer.
                     */
                    name(value?:  string): any;
                    /**
                     * Gets or sets whether to show the header of the item slicer.
                     * @param {boolean} value The show header setting of the item slicer.
                     * @returns {boolean | GC.Spread.Sheets.Slicers.ItemSlicer} If no value is set, returns whether to show the header of the item slicer; otherwise, returns the item slicer.
                     */
                    showHeader(value?:  boolean): any;
                    /**
                     * Gets or sets whether to show the no data items of the item slicer.
                     * @param {boolean} value The show no data items setting of the slicer.
                     * @returns {boolean | GC.Spread.Sheets.Slicers.ItemSlicer} If no value is set, returns whether to show the no data items of the item slicer; otherwise, returns the item slicer.
                     */
                    showNoDataItems(value?:  boolean): any;
                    /**
                     * Gets or sets whether to show the no data items last.
                     * @param {boolean} value The show no data items in last setting of the slicer.
                     * @returns {boolean | GC.Spread.Sheets.Slicers.ItemSlicer} If no value is set, returns whether to show the no data items last; otherwise, returns the item slicer.
                     */
                    showNoDataItemsInLast(value?:  boolean): any;
                    /**
                     * Gets or sets the sort state of the item slicer.
                     * @param {GC.Spread.Sheets.SortState} value The sort state of the item slicer.
                     * @returns {GC.Spread.Sheets.SortState | GC.Spread.Sheets.Slicers.ItemSlicer} If no value is set, returns the sort state of the item slicer; otherwise, returns the item slicer.
                     */
                    sortState(value?:  GC.Spread.Sheets.SortState): any;
                    /**
                     * Gets or sets the style of the item slicer.
                     * @param {any} value The style of the item slicer.
                     * @returns {any | GC.Spread.Sheets.Slicers.ItemSlicer} If no value is set, returns The style of the item slicer; otherwise, returns the item slicer.
                     * @example
                     * The style is json data, it's json schema is as follows:
                     * {
                     *      "$schema" : "http://json-schema.org/draft-04/schema#",
                     *      "title" : "style",
                     *      "type" : "object",
                     *      "properties" : {
                     *          "wholeSlicerStyle" : {
                     *              "$ref" : "#/definitions/StyleInfo"
                     *          },
                     *          "headerStyle" : {
                     *              "$ref" : "#/definitions/StyleInfo"
                     *          },
                     *          "selectedItemWithDataStyle" : {
                     *              "$ref" : "#/definitions/StyleInfo"
                     *          },
                     *          "selectedItemWithNoDataStyle" : {
                     *              "$ref" : "#/definitions/StyleInfo"
                     *          },
                     *          "unSelectedItemWithDataStyle" : {
                     *              "$ref" : "#/definitions/StyleInfo"
                     *          },
                     *          "unSelectedItemWithNoDataStyle" : {
                     *              "$ref" : "#/definitions/StyleInfo"
                     *          },
                     *          "hoveredSelectedItemWithDataStyle" : {
                     *              "$ref" : "#/definitions/StyleInfo"
                     *          },
                     *          "hoveredSelectedItemWithNoDataStyle" : {
                     *              "$ref" : "#/definitions/StyleInfo"
                     *          },
                     *          "hoveredUnSelectedItemWithDataStyle" : {
                     *              "$ref" : "#/definitions/StyleInfo"
                     *          },
                     *          "hoveredUnSelectedItemWithNoDataStyle" : {
                     *              "$ref" : "#/definitions/StyleInfo"
                     *          }
                     *      },
                     *      "definitions" : {
                     *          "StyleInfo" : {
                     *              "type" : "object",
                     *              "properties" : {
                     *                      "backColor" : {
                     *                              "type" : "string"
                     *                      },
                     *                      "foreColor" : {
                     *                              "type" : "string"
                     *                      },
                     *                      "font" : {
                     *                              "type" : "string"
                     *                      },
                     *                      "borderLeft" : {
                     *                              "$ref" : "#/definitions/SlicerBorder"
                     *                      },
                     *                      "borderTop" : {
                     *                              "$ref" : "#/definitions/SlicerBorder"
                     *                      },
                     *                      "borderRight" : {
                     *                              "$ref" : "#/definitions/SlicerBorder"
                     *                      },
                     *                      "borderBottom" : {
                     *                              "$ref" : "#/definitions/SlicerBorder"
                     *                      },
                     *                  "textDecoration":{
                     *                      "type" : "string"
                     *                  }
                     *              }
                     *          },
                     *          "SlicerBorder":{
                     *              "type":"object",
                     *              "properties":{
                     *                  "borderWidth":{
                     *                          "type":"number"
                     *                  },
                     *                  "borderStyle":{
                     *                          "type":"string"
                     *                  },
                     *                  "borderColor":{
                     *                          "type":"string"
                     *                  }
                     *           }
                     *          }
                     *      }
                     *  }
                     */
                    style(value?:  any): any;
                    /**
                     * Gets or sets whether to visually distinguish the items with no data.
                     * @param {boolean} value The setting for items with no data.
                     * @returns {boolean | GC.Spread.Sheets.Slicers.ItemSlicer} If no value is set, returns whether to visually distinguish the items with no data; otherwise, returns the item slicer.
                     */
                    visuallyNoDataItems(value?:  boolean): any;
                    /**
                     * Gets or sets the width of the item slicer.
                     * @param {number} value The width of the item slicer.
                     * @returns {number | GC.Spread.Sheets.Slicers.ItemSlicer} If no value is set, returns the width of the item slicer; otherwise, returns the item slicer.
                     */
                    width(value?:  number): any;
                }

                export class Slicer extends GC.Spread.Sheets.FloatingObjects.FloatingObject{
                    /**
                     * Represents a slicer.
                     * @class GC.Spread.Sheets.Slicer
                     * @param {string} name The slicer name.
                     * @param {GC.Spread.Sheets.Tables.Table} table The table that relates to the slicer.
                     * @param {string} columnName The name of the table's column.
                     */
                    constructor(name:  string,  table:  GC.Spread.Sheets.Tables.Table,  columnName:  string);
                    /**
                     * Gets or sets the caption name of the slicer.
                     * @param {string} value The caption name of the slicer.
                     * @returns {string | GC.Spread.Sheets.Slicers.Slicer} If no value is set, returns the caption name of the slicer; otherwise, returns the slicer.
                     */
                    captionName(value?:  string): any;
                    /**
                     * Gets or sets the column count for the slicer.
                     * @param {number} value The column count of the slicer.
                     * @returns {number | GC.Spread.Sheets.Slicers.Slicer} If no value is set, returns the column count for the slicer; otherwise, returns the slicer.
                     */
                    columnCount(value?:  number): any;
                    /**
                     * Gets or sets whether to disable resizing and moving the slicer.
                     * @param {boolean} value The setting for whether to disable resizing and moving the slicer.
                     * @returns {boolean | GC.Spread.Sheets.Slicers.Slicer} If no value is set, returns whether to disable resizing and moving the slicer; otherwise, returns the slicer.
                     */
                    disableResizingAndMoving(value?:  boolean): any;
                    /**
                     * Gets or sets the item height for the slicer.
                     * @param {number} value The item height of the slicer.
                     * @returns {number | GC.Spread.Sheets.Slicers.Slicer} If no value is set, returns the item height for the slicer; otherwise, returns the slicer.
                     */
                    itemHeight(value?:  number): any;
                    /**
                     * Gets or sets the name of the slicer.
                     * @param {string} value The name of the slicer.
                     * @returns {string | GC.Spread.Sheets.Slicers.Slicer} If no value is set, returns the name of the slicer; otherwise, returns the slicer.
                     */
                    name(value?:  string): any;
                    /**
                     * Gets or sets whether to show the slicer header.
                     * @param {boolean} value The show header setting of the slicer.
                     * @returns {boolean | GC.Spread.Sheets.Slicers.Slicer} If no value is set, returns whether to show the slicer header; otherwise, returns the slicer.
                     */
                    showHeader(value?:  boolean): any;
                    /**
                     * Gets or sets whether to show the no data items of the slicer.
                     * @param {boolean} value The show no data items setting of the slicer.
                     * @returns {boolean | GC.Spread.Sheets.Slicers.Slicer} If no value is set, returns whether to show the no data items of the slicer; otherwise, returns the slicer.
                     */
                    showNoDataItems(value?:  boolean): any;
                    /**
                     * Gets or sets whether to show the no data items last.
                     * @param {boolean} value The show no data items last setting of the slicer.
                     * @returns {boolean | GC.Spread.Sheets.Slicers.Slicer} If no value is set, returns whether to show the no data items last; otherwise, returns the slicer.
                     */
                    showNoDataItemsInLast(value?:  boolean): any;
                    /**
                     * Gets or sets the sort state of the slicer.
                     * @param {GC.Spread.Sheets.SortState} value The sort state of the slicer.
                     * @returns {GC.Spread.Sheets.SortState | GC.Spread.Sheets.Slicers.Slicer} If no value is set, returns the sort state of the slicer; otherwise, returns the slicer.
                     */
                    sortState(value?:  GC.Spread.Sheets.SortState): any;
                    /**
                     * Gets or sets the style of the slicer.
                     * @param {GC.Spread.Sheets.Slicers.SlicerStyle} value The style of the slicer.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle | GC.Spread.Sheets.Slicers.Slicer} If no value is set, returns The style of the slicer; otherwise, returns the slicer.
                     */
                    style(value?:  GC.Spread.Sheets.Slicers.SlicerStyle): any;
                    /**
                     * Gets or sets whether to visually distinguish the items with no data.
                     * @param {boolean} value The setting for items with no data.
                     * @returns {boolean | GC.Spread.Sheets.Slicers.Slicer} If no value is set, returns whether to visually distinguish the items with no data; otherwise, returns the slicer.
                     */
                    visuallyNoDataItems(value?:  boolean): any;
                }

                export class SlicerBorder{
                    /**
                     * Represents the slicer border.
                     * @param {number} borderWidth The border width.
                     * @param {string} borderStyle The border style.
                     * @param {string} borderColor The border color.
                     * @class
                     */
                    constructor(borderWidth:  number,  borderStyle:  string,  borderColor:  string);
                    /**
                     * Gets or sets the border color.
                     * @param {string} value The border color.
                     * @returns {string | GC.Spread.Sheets.Slicers.SlicerBorder}  If no value is set, returns the border color; otherwise, returns the slicer border.
                     */
                    borderColor(value?:  string): any;
                    /**
                     * Gets or sets the border style.
                     * @param {string} value The border style.
                     * @returns {string | GC.Spread.Sheets.Slicers.SlicerBorder}  If no value is set, returns the border style; otherwise, returns the slicer border.
                     */
                    borderStyle(value?:  string): any;
                    /**
                     * Gets or sets the border width.
                     * @param {number} value The border width.
                     * @returns {number | GC.Spread.Sheets.Slicers.SlicerBorder}  If no value is set, returns the border width; otherwise, returns the slicer border.
                     */
                    borderWidth(value?:  number): any;
                }

                export class SlicerCollection{
                    /**
                     * Represents a slicer manager that managers all slicers in a sheet.
                     * @class
                     * @param {GC.Spread.Sheets.Worksheet} sheet The worksheet.
                     */
                    constructor(sheet:  GC.Spread.Sheets.Worksheet);
                    /**
                     * Adds a slicer to the sheet.
                     * @param {string} name The name of the slicer.
                     * @param {string} tableName The name of the table that relates to the slicer.
                     * @param {string} columnName The name of the table column that relates to the slicer.
                     * @param {GC.Spread.Sheets.Slicers.SlicerStyle} style The style of the slicer.
                     * @return {GC.Spread.Sheets.Slicers.Slicer} The slicer that has been added to the sheet.
                     */
                    add(name:  string,  tableName:  string,  columnName:  string,  style:  GC.Spread.Sheets.Slicers.SlicerStyle): GC.Spread.Sheets.Slicers.Slicer;
                    /**
                     * Gets all of the slicers in the sheet with the indicated table name and column name.
                     * @param tableName {string} The name of the table.
                     * @param columnName {string} The name of the column.
                     * @returns {Array} The slicer collection.
                     */
                    all(tableName:  string,  columnName:  string): GC.Spread.Sheets.Slicers.Slicer[];
                    /**
                     * Removes all of the slicers from the sheet.
                     */
                    clear(): any;
                    /**
                     * Gets a slicer in the sheet by the name.
                     * @param {string} name The name of the slicer.
                     * @returns {GC.Spread.Sheets.Slicers.Slicer} The slicer that has the indicated name.
                     */
                    get(name:  string): GC.Spread.Sheets.Slicers.Slicer;
                    /**
                     * Removes a slicer from the sheet using the indicated slicer name.
                     * @param {string} name The name of the slicer.
                     */
                    remove(name:  string): void;
                }

                export class SlicerStyle{
                    /**
                     * Represents the slicer style settings.
                     * @class
                     */
                    constructor();
                    /**
                     * Gets or sets the style of the slicer header.
                     * @param {GC.Spread.Sheets.Slicers.SlicerStyleInfo} value The style of the slicer header.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyleInfo | GC.Spread.Sheets.Slicers.SlicerStyle} If no value is set, returns the style of the slicer header; otherwise, returns the slicer style.
                     */
                    headerStyle(value?:  GC.Spread.Sheets.Slicers.SlicerStyleInfo): any;
                    /**
                     * Gets or sets the style of the hovered selected item with data.
                     * @param {GC.Spread.Sheets.Slicers.SlicerStyleInfo} value The style of the hovered selected item with data.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyleInfo | GC.Spread.Sheets.Slicers.SlicerStyle} If no value is set, returns the style of the hovered selected item with data; otherwise, returns the slicer style.
                     */
                    hoveredSelectedItemWithDataStyle(value?:  GC.Spread.Sheets.Slicers.SlicerStyleInfo): any;
                    /**
                     * Gets or sets the style of the hovered selected item with no data.
                     * @param {GC.Spread.Sheets.Slicers.SlicerStyleInfo} value The style of the hovered selected item with no data.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyleInfo | GC.Spread.Sheets.Slicers.SlicerStyle} If no value is set, returns the style of the hovered selected item with no data; otherwise, returns the slicer style.
                     */
                    hoveredSelectedItemWithNoDataStyle(value?:  GC.Spread.Sheets.Slicers.SlicerStyleInfo): any;
                    /**
                     * Gets or sets the style of the hovered unselected item with data.
                     * @param {GC.Spread.Sheets.Slicers.SlicerStyleInfo} value The style of the hovered unselected item with data.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyleInfo | GC.Spread.Sheets.Slicers.SlicerStyle} If no value is set, returns the style of the hovered unselected item with data; otherwise, returns the slicer style.
                     */
                    hoveredUnSelectedItemWithDataStyle(value?:  GC.Spread.Sheets.Slicers.SlicerStyleInfo): any;
                    /**
                     * Gets or sets the style of the hovered unselected item with no data.
                     * @param {GC.Spread.Sheets.Slicers.SlicerStyleInfo} value The style of the hovered unselected item with no data.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyleInfo | GC.Spread.Sheets.Slicers.SlicerStyle} If no value is set, returns the style of the hovered unselected item with no data; otherwise, returns the slicer style.
                     */
                    hoveredUnSelectedItemWithNoDataStyle(value?:  GC.Spread.Sheets.Slicers.SlicerStyleInfo): any;
                    /**
                     * Gets or sets the name of the style.
                     * @param {string} value The slicer style name.
                     * @returns {string | GC.Spread.Sheets.Slicers.SlicerStyle} If no value is set, returns the name of the style; otherwise, returns the slicer style.
                     */
                    name(value?:  string): any;
                    /**
                     * Gets or sets the style of the selected item with data.
                     * @param {GC.Spread.Sheets.Slicers.SlicerStyleInfo} value The style of the selected item with data.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyleInfo | GC.Spread.Sheets.Slicers.SlicerStyle} If no value is set, returns the style of the selected item with data; otherwise, returns the slicer style.
                     */
                    selectedItemWithDataStyle(value?:  GC.Spread.Sheets.Slicers.SlicerStyleInfo): any;
                    /**
                     * Gets or sets the style of the selected item with no data.
                     * @param {GC.Spread.Sheets.Slicers.SlicerStyleInfo} value The style of the selected item with no data.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyleInfo | GC.Spread.Sheets.Slicers.SlicerStyle} If no value is set, returns the style of the selected item with no data; otherwise, returns the slicer style.
                     */
                    selectedItemWithNoDataStyle(value?:  GC.Spread.Sheets.Slicers.SlicerStyleInfo): any;
                    /**
                     * Gets or sets the style of the unselected item with data.
                     * @param {GC.Spread.Sheets.Slicers.SlicerStyleInfo} value The style of the unselected item with data.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyleInfo | GC.Spread.Sheets.Slicers.SlicerStyle} If no value is set, returns the style of the unselected item with data; otherwise, returns the slicer style.
                     */
                    unSelectedItemWithDataStyle(value?:  GC.Spread.Sheets.Slicers.SlicerStyleInfo): any;
                    /**
                     * Gets or sets the style of the unselected item with no data.
                     * @param {GC.Spread.Sheets.Slicers.SlicerStyleInfo} value The style of the unselected item with no data.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyleInfo | GC.Spread.Sheets.Slicers.SlicerStyle} If no value is set, returns the style of the unselected item with no data; otherwise, returns the slicer style.
                     */
                    unSelectedItemWithNoDataStyle(value?:  GC.Spread.Sheets.Slicers.SlicerStyleInfo): any;
                    /**
                     * Gets or sets the style of the whole slicer.
                     * @param {GC.Spread.Sheets.Slicers.SlicerStyleInfo} value The style of the whole slicer.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyleInfo | GC.Spread.Sheets.Slicers.SlicerStyle} If no value is set, returns the style of the whole slicer; otherwise, returns the slicer style.
                     */
                    wholeSlicerStyle(value?:  GC.Spread.Sheets.Slicers.SlicerStyleInfo): any;
                }

                export class SlicerStyleInfo{
                    /**
                     * Represents slicer style information.
                     * @class
                     * @param {string} backColor The background color of the style information.
                     * @param {string} foreColor The foreground color of the style information.
                     * @param {string} font The font of the style information.
                     * @param {GC.Spread.Sheets.Slicers.SlicerBorder} borderLeft The left border of the style information.
                     * @param {GC.Spread.Sheets.Slicers.SlicerBorder} borderTop The top border of the slicer information.
                     * @param {GC.Spread.Sheets.Slicers.SlicerBorder} borderRight The right border of the style information.
                     * @param {GC.Spread.Sheets.Slicers.SlicerBorder} borderBottom The bottom border of the style information.
                     * @param {GC.Spread.Sheets.TextDecorationType} textDecoration The text decoration of the style information.
                     */
                    constructor(backColor?:  string,  foreColor?:  string,  font?:  string,  borderLeft?:  GC.Spread.Sheets.Slicers.SlicerBorder,  borderTop?:  GC.Spread.Sheets.Slicers.SlicerBorder,  borderRight?:  GC.Spread.Sheets.Slicers.SlicerBorder,  borderBottom?:  GC.Spread.Sheets.Slicers.SlicerBorder,  textDecoration?:  GC.Spread.Sheets.TextDecorationType);
                    /**
                     * Gets or sets the background color of the style information.
                     * @param {string} value The background color of the style information.
                     * @returns {string | GC.Spread.Sheets.Slicers.SlicerStyleInfo}  If no value is set, returns the background color of the style information; otherwise, returns the slicer style information.
                     */
                    backColor(value?:  string): any;
                    /**
                     * Gets or sets the bottom border of the style information.
                     * @param {GC.Spread.Sheets.Slicers.SlicerBorder} value The bottom border of the style information.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerBorder | GC.Spread.Sheets.Slicers.SlicerStyleInfo}  If no value is set, returns the bottom border of the style information; otherwise, returns the slicer style information.
                     */
                    borderBottom(value?:  GC.Spread.Sheets.Slicers.SlicerBorder): any;
                    /**
                     * Gets or sets the left border of the style information.
                     * @param {GC.Spread.Sheets.Slicers.SlicerBorder} value The left border of the style information.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerBorder | GC.Spread.Sheets.Slicers.SlicerStyleInfo}  If no value is set, returns the left border of the style information; otherwise, returns the slicer style information.
                     */
                    borderLeft(value?:  GC.Spread.Sheets.Slicers.SlicerBorder): any;
                    /**
                     * Gets or sets the right border of the style information.
                     * @param {GC.Spread.Sheets.Slicers.SlicerBorder} value The right border of the style information.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerBorder | GC.Spread.Sheets.Slicers.SlicerStyleInfo}  If no value is set, returns the right border of the style information; otherwise, returns the slicer style information.
                     */
                    borderRight(value?:  GC.Spread.Sheets.Slicers.SlicerBorder): any;
                    /**
                     * Gets or sets the top border of the style information.
                     * @param {GC.Spread.Sheets.Slicers.SlicerBorder} value The top border of the style information.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerBorder | GC.Spread.Sheets.Slicers.SlicerStyleInfo}  If no value is set, returns the top border of the style information; otherwise, returns the slicer style information.
                     */
                    borderTop(value?:  GC.Spread.Sheets.Slicers.SlicerBorder): any;
                    /**
                     * Gets or sets the font of the style information.
                     * @param {string} value The font of the style information.
                     * @returns {string | GC.Spread.Sheets.Slicers.SlicerStyleInfo}  If no value is set, returns the font of the style information; otherwise, returns the slicer style information.
                     */
                    font(value?:  string): any;
                    /**
                     * Gets or sets the foreground color of the style information.
                     * @param {string} value The foreground color of the style information.
                     * @returns {string | GC.Spread.Sheets.Slicers.SlicerStyleInfo}  If no value is set, returns the foreground color of the style information; otherwise, returns the slicer style information.
                     */
                    foreColor(value?:  string): any;
                    /**
                     * Sets every border of the style information.
                     * @param {GC.Spread.Sheets.Slicers.SlicerBorder} value The border setting.
                     */
                    setBorders(value:  GC.Spread.Sheets.Slicers.SlicerBorder): void;
                    /**
                     * Gets or sets the text decoration of the style information.
                     * @param {GC.Spread.Sheets.TextDecorationType} value The text decoration of the style information.
                     * @returns {GC.Spread.Sheets.TextDecorationType | GC.Spread.Sheets.Slicers.SlicerStyleInfo}  If no value is set, returns the text decoration of the style information; otherwise, returns the slicer style information.
                     */
                    textDecoration(value?:  GC.Spread.Sheets.TextDecorationType): any;
                }

                export class SlicerStyles{
                    /**
                     * Represents a built-in slicer style collection.
                     * @class
                     */
                    constructor();
                    /**
                     * Gets the dark1 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static dark1(): GC.Spread.Sheets.Slicers.SlicerStyle;
                    /**
                     * Gets the dark2 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static dark2(): GC.Spread.Sheets.Slicers.SlicerStyle;
                    /**
                     * Gets the dark3 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static dark3(): GC.Spread.Sheets.Slicers.SlicerStyle;
                    /**
                     * Gets the dark4 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static dark4(): GC.Spread.Sheets.Slicers.SlicerStyle;
                    /**
                     * Gets the dark5 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static dark5(): GC.Spread.Sheets.Slicers.SlicerStyle;
                    /**
                     * Gets the dark6 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static dark6(): GC.Spread.Sheets.Slicers.SlicerStyle;
                    /**
                     * Gets the light1 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static light1(): GC.Spread.Sheets.Slicers.SlicerStyle;
                    /**
                     * Gets the light2 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static light2(): GC.Spread.Sheets.Slicers.SlicerStyle;
                    /**
                     * Gets the light3 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static light3(): GC.Spread.Sheets.Slicers.SlicerStyle;
                    /**
                     * Gets the light4 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static light4(): GC.Spread.Sheets.Slicers.SlicerStyle;
                    /**
                     * Gets the light5 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static light5(): GC.Spread.Sheets.Slicers.SlicerStyle;
                    /**
                     * Gets the light6 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static light6(): GC.Spread.Sheets.Slicers.SlicerStyle;
                    /**
                     * Gets the other1 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static other1(): GC.Spread.Sheets.Slicers.SlicerStyle;
                    /**
                     * Gets the other2 style.
                     * @returns {GC.Spread.Sheets.Slicers.SlicerStyle}
                     */
                    static other2(): GC.Spread.Sheets.Slicers.SlicerStyle;
                }

                export class TableSlicerData extends GC.Spread.Slicers.GeneralSlicerData{
                    /**
                     * Represents table slicer data.
                     * @extends GC.Spread.Slicers.GeneralSlicerData
                     * @class GC.Spread.Sheets.TableSlicerData
                     * @param {GC.Spread.Sheets.Tables.Table} table The table.
                     */
                    constructor(table:  GC.Spread.Sheets.Tables.Table);
                    /**
                     * Filters the data that corresponds to the specified column name and exclusive data indexes.
                     * @param {string} columnName The column name.
                     * @param {object} conditional The filter conditional.<br />
                     * conditional.exclusiveRowIndexes: number array type, visible exclusive row indexes<br />
                     * conditional.ranges: {min:number, max:number} array type, visible ranges.
                     * @param {boolean} isPreview Indicates whether filter is in preview mode.
                     */
                    doFilter(columnName:  string,  conditional:  GC.Spread.Slicers.ISlicerConditional,  isPreview?:  boolean): void;
                    /**
                     * Unfilters the data that corresponds to the specified column name.
                     * @param {string} columnName The column name.
                     */
                    doUnfilter(columnName:  string): void;
                    /**
                     * Gets the slicer data of the table.
                     * @returns {GC.Spread.Sheets.Slicers.TableSlicerData} The slicer data of the table.
                     */
                    getSlicerData(): GC.Spread.Sheets.Slicers.TableSlicerData;
                    /**
                     * Gets the table of the table slicer data.
                     * @returns {GC.Spread.Sheets.Tables.Table} The table of the table slicer data.
                     */
                    getTable(): GC.Spread.Sheets.Tables.Table;
                    /**
                     * Refreshes the table slicer data.
                     */
                    refresh(): void;
                }
            }

            module Sparklines{

                export interface ISparklineSetting{
                    axisColor: string;
                    firstMarkerColor: string;
                    highMarkerColor: string;
                    lastMarkerColor: string;
                    lowMarkerColor: string;
                    markersColor: string;
                    negativeColor: string;
                    seriesColor: string;
                    displayEmptyCellsAs: EmptyValueStyle;
                    rightToLeft: boolean;
                    displayHidden: boolean;
                    displayXAxis: boolean;
                    showFirst: boolean;
                    showHigh: boolean;
                    showLast: boolean;
                    showLow: boolean;
                    showNegative: boolean;
                    showMarkers: boolean;
                    manualMax: number;
                    manualMin: number;
                    maxAxisType: SparklineAxisMinMax;
                    minAxisType: SparklineAxisMinMax;
                    groupMaxValue: number;
                    groupMinValue: number;
                    lineWeight: number;
                }

                /**
                 * Represents the orientation of the range.
                 * @enum {number}
                 */
                export enum DataOrientation{
                    /** Specifies the vertical orientation.
                     * @type {number}
                     */
                    vertical= 0,
                    /** Specifies the horizontal orientation.
                     * @type {number}
                     */
                    horizontal= 1
                }

                /**
                 * Specifies how to show an empty value from a data series in the chart.
                 * @enum {number}
                 */
                export enum EmptyValueStyle{
                    /** Leaves gaps for empty values in a data series, which results in a segmented line.
                     * @type {number}
                     */
                    gaps= 0,
                    /** Handles empty values in a data series as zero values, so that the line drops to zero for zero-value data points.
                     * @type {number}
                     */
                    zero= 1,
                    /** Fills gaps with a connecting element instead of leaving gaps for empty values in a data series.
                     * @type {number}
                     */
                    connect= 2
                }

                /**
                 * An enumeration that specifies information about how the vertical axis minimum or maximum is computed for this sparkline group.
                 * @enum {number}
                 */
                export enum SparklineAxisMinMax{
                    /** Specifies that the vertical axis minimum or maximum for each sparkline in this sparkline group is calculated automatically such that the data point with the minimum or maximum value can be displayed in the plot area.
                     * @type {number}
                     */
                    individual= 0,
                    /** Specifies that the vertical axis minimum or maximum is shared across all sparklines in this sparkline group and is calculated automatically such that the data point with the minimum or maximum value can be displayed in the plot area.
                     * @type {number}
                     */
                    group= 1,
                    /** Specifies that the vertical axis minimum or maximum for each sparkline in this sparkline group is specified by the manualMin attribute or the manualMax attribute of the sparkline group.
                     * @type {number}
                     */
                    custom= 2
                }

                /**
                 * Represents the sparkline type.
                 * @enum {number}
                 */
                export enum SparklineType{
                    /** Specifies the line sparkline.
                     * @type {number}
                     */
                    line= 0,
                    /** Specifies the column sparkline.
                     * @type {number}
                     */
                    column= 1,
                    /** Specifies the win-loss sparkline.
                     * @type {number}
                     */
                    winloss= 2
                }


                export class AreaSparkline extends SparklineEx{
                    /**
                     * Represents the class for the area sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class BoxPlotSparkline extends SparklineEx{
                    /**
                     * Represents the class for the box plot sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class BulletSparkline extends SparklineEx{
                    /**
                     * Represents the class for the bullet sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class CascadeSparkline extends SparklineEx{
                    /**
                     * Represents the class for the cascade sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class ColumnSparkline extends SparklineEx{
                    /**
                     * Represents the class for the column sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class HBarSparkline extends SparklineEx{
                    /**
                     * Represents the class for the Hbar sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class LineSparkline extends SparklineEx{
                    /**
                     * Represents the class for the line sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class MonthSparkline extends SparklineEx{
                    /**
                     * Represents the class for the month sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class ParetoSparkline extends SparklineEx{
                    /**
                     * Represents the class for the pareto sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class PieSparkline extends SparklineEx{
                    /**
                     * Represents the class for the pie sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class ScatterSparkline extends SparklineEx{
                    /**
                     * Represents the class for the scatter sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class Sparkline{
                    /**
                     * Represents a Sparkline class.
                     * @param {number} row The row index.
                     * @param {number} column The column index.
                     * @param {GC.Spread.Sheets.Range} dataReference The data range to which the sparkline refers.
                     * @param {GC.Spread.Sheets.Sparklines.DataOrientation} dataOrientation The orientation of the data range.
                     * @param {GC.Spread.Sheets.Sparklines.SparklineType} type The type of sparkline.
                     * @param {GC.Spread.Sheets.Sparklines.SparklineSetting} setting The setting of the sparkline.
                     * @class
                     */
                    constructor(row?:  number,  column?:  number,  dataReference?:  GC.Spread.Sheets.Range,  dataOrientation?:  GC.Spread.Sheets.Sparklines.DataOrientation,  type?:  GC.Spread.Sheets.Sparklines.SparklineType,  setting?:  GC.Spread.Sheets.Sparklines.SparklineSetting);
                    /** Gets the column index.
                     * @type {number}
                     */
                    column: number;
                    /** Gets the row index.
                     * @type {number}
                     */
                    row: number;
                    /**
                     * Clones a sparkline.
                     * @returns {GC.Spread.Sheets.Sparklines.Sparkline} The cloned sparkline.
                     */
                    clone(): GC.Spread.Sheets.Sparklines.Sparkline;
                    /**
                     * Gets or sets the data object.
                     * @param {GC.Spread.Sheets.Range} value The sparkline data.
                     * @returns {GC.Spread.Sheets.Range | GC.Spread.Sheets.Sparklines.Sparkline} If no value is set, returns the data object; otherwise, returns the sparkline.
                     */
                    data(value?:  GC.Spread.Sheets.Range): any;
                    /**
                     * Gets or sets the data orientation.
                     * @param {GC.Spread.Sheets.Sparklines.DataOrientation} value The sparkline data orientation.
                     * @returns {GC.Spread.Sheets.Sparklines.DataOrientation | GC.Spread.Sheets.Sparklines.Sparkline} If no value is set, returns the sparkline data orientation; otherwise, returns the sparkline.
                     */
                    dataOrientation(value?:  GC.Spread.Sheets.Sparklines.DataOrientation): any;
                    /**
                     * Gets or sets the date axis data object.
                     * @param {GC.Spread.Sheets.Range} value The sparkline date axis data.
                     * @returns {GC.Spread.Sheets.Range | GC.Spread.Sheets.Sparklines.Sparkline} If no value is set, returns the sparkline date axis data; otherwise, returns the sparkline.
                     */
                    dateAxisData(value?:  GC.Spread.Sheets.Range): any;
                    /**
                     * Gets or sets the date axis orientation.
                     * @param {GC.Spread.Sheets.Sparklines.DataOrientation} value The sparkline date axis orientation.
                     * @returns {GC.Spread.Sheets.Sparklines.DataOrientation | GC.Spread.Sheets.Sparklines.Sparkline} If no value is set, returns the sparkline date axis orientation; otherwise, returns the sparkline.
                     */
                    dateAxisOrientation(value?:  GC.Spread.Sheets.Sparklines.DataOrientation): any;
                    /**
                     * Gets or sets a value that indicates whether to display the date axis.
                     * @param {boolean} value Whether to display the date axis.
                     * @returns {boolean | GC.Spread.Sheets.Sparklines.Sparkline} If no value is set, returns whether to display the date axis; otherwise, returns the sparkline.
                     */
                    displayDateAxis(value?:  boolean): any;
                    /**
                     * Gets or sets the sparkline group.
                     * @param {GC.Spread.Sheets.Sparklines.SparklineGroup} value The sparkline group.
                     * @returns {GC.Spread.Sheets.Sparklines.SparklineGroup | GC.Spread.Sheets.Sparklines.Sparkline} If no value is set, returns the sparkline group; otherwise, returns the sparkline.
                     */
                    group(value?:  GC.Spread.Sheets.Sparklines.SparklineGroup): any;
                    /**
                     * Paints the sparkline in the specified area.
                     * @param {CanvasRenderingContext2D} ctx The canvas's two-dimensional context.
                     * @param {number} x <i>x</i>-coordinate relative to the canvas.
                     * @param {number} y <i>y</i>-coordinate relative to the canvas.
                     * @param {number} w The width of the cell that contains the sparkline.
                     * @param {number} h The height of the cell that contains the sparkline.
                     */
                    paintSparkline(ctx:  CanvasRenderingContext2D,  x:  number,  y:  number,  w:  number,  h:  number): void;
                    /**
                     * Gets or sets the sparkline setting of the cell.
                     * @param {GC.Spread.Sheets.Sparklines.SparklineSetting} value The sparkline setting.
                     * @returns {GC.Spread.Sheets.Sparklines.SparklineSetting  | GC.Spread.Sheets.Sparklines.Sparkline} If no value is set, returns the sparkline setting; otherwise, returns the sparkline.
                     */
                    setting(value?:  GC.Spread.Sheets.Sparklines.SparklineSetting): any;
                    /**
                     * Gets or sets the type of the sparkline.
                     * @param {GC.Spread.Sheets.Sparklines.SparklineType} value The sparkline type.
                     * @returns {GC.Spread.Sheets.Sparklines.SparklineType | GC.Spread.Sheets.Sparklines.Sparkline} If no value is set, returns the sparkline type; otherwise, returns the sparkline.
                     */
                    sparklineType(value?:  GC.Spread.Sheets.Sparklines.SparklineType): any;
                }

                export class SparklineEx{
                    /**
                     * Represents the base class for the other SparklineEx classes.
                     * @class
                     */
                    constructor();
                    /**
                     * Represents the type name string used for supporting serialization.
                     * @type {string}
                     */
                    typeName: string;
                    /**
                     * Creates a custom function used to provide data and settings for SparklineEx.
                     * @returns {GC.Spread.CalcEngine.Functions.Function} The created custom function.
                     */
                    createFunction(): GC.Spread.CalcEngine.Functions.Function;
                    /**
                     * Loads the object state from the specified JSON string.
                     * @param {Object} settings The sparklineEx data from deserialization.
                     */
                    fromJSON(settings:  Object): void;
                    /**
                     * Gets the name of SparklineEx.
                     * @returns {string} The SparklineEx's name.
                     */
                    name(): string;
                    /**
                     * Paints the SparklineEx on the canvas.
                     * @param {CanvasRenderingContext2D} context The canvas's two-dimensional context.
                     * @param {object} value The value evaluated by the custom function.
                     * @param {number} x <i>x</i>-coordinate relative to the canvas.
                     * @param {number} y <i>y</i>-coordinate relative to the canvas.
                     * @param {number} width The cell's width.
                     * @param {number} height The cell's height.
                     */
                    paint(context:  CanvasRenderingContext2D,  value:  any,  x:  number,  y:  number,  width:  number,  height:  number): void;
                    /**
                     * Saves the object state to a JSON string.
                     * @returns {Object} The sparklineEx data.
                     */
                    toJSON(): Object;
                }

                export class SparklineGroup{
                    /**
                     * Represents a sparkline group.
                     * @param {GC.Spread.Sheets.Sparklines.SparklineType} type The type of sparkline.
                     * @param {GC.Spread.Sheets.Sparklines.SparklineSetting} setting The setting of the sparkline group.
                     * @class
                     */
                    constructor(type:  GC.Spread.Sheets.Sparklines.SparklineType,  setting:  GC.Spread.Sheets.Sparklines.SparklineSetting);
                    /** Indicates the sparkline settings.
                     * @type {GC.Spread.Sheets.Sparklines.SparklineSetting}
                     */
                    setting: GC.Spread.Sheets.Sparklines.SparklineSetting;
                    /** Indicates the sparkline type.
                     * @type {GC.Spread.Sheets.Sparklines.SparklineType}
                     */
                    sparklineType: GC.Spread.Sheets.Sparklines.SparklineType;
                    /**
                     * Adds a sparkline to the group.
                     * @param {GC.Spread.Sheets.Sparklines.Sparkline} item The sparkline item.
                     */
                    add(item:  GC.Spread.Sheets.Sparklines.Sparkline): void;
                    /**
                     * Clones the current sparkline group.
                     * @returns {GC.Spread.Sheets.Sparklines.SparklineGroup} The cloned sparkline group.
                     */
                    clone(): GC.Spread.Sheets.Sparklines.SparklineGroup;
                    /**
                     * Determines whether the group contains a specific value.
                     * @param {GC.Spread.Sheets.Sparklines.Sparkline} item The object to locate in the group.
                     * @returns {boolean} <c>true</c> if the item is found in the group; otherwise, <c>false</c>.
                     */
                    contains(item:  GC.Spread.Sheets.Sparklines.Sparkline): boolean;
                    /**
                     * Represents the count of the sparkline group innerlist.
                     * @returns {number} The sparkline count in the group.
                     */
                    count(): number;
                    /**
                     * Represents the date axis data.
                     * @param {GC.Spread.Sheets.Range} value The date axis data.
                     * @returns {GC.Spread.Sheets.Range | undefined} If no value is set, returns the date axis data; otherwise, returns undefined.
                     */
                    dateAxisData(value?:  GC.Spread.Sheets.Range): any;
                    /**
                     * Represents the date axis orientation.
                     * @param {GC.Spread.Sheets.Sparklines.DataOrientation} value The date axis orientation.
                     * @returns {GC.Spread.Sheets.Sparklines.DataOrientation | undefined} If no value is set, returns the date axis orientation; otherwise, returns undefined.
                     */
                    dateAxisOrientation(value:  GC.Spread.Sheets.Sparklines.DataOrientation): any;
                    /**
                     * Removes the first occurrence of a specific object from the group.
                     * @param {GC.Spread.Sheets.Sparklines.Sparkline} item The sparkline item.
                     * @returns {Array} The GC.Spread.Sheets.Sparklines.Sparkline array.
                     */
                    remove(item:  GC.Spread.Sheets.Sparklines.Sparkline): GC.Spread.Sheets.Sparklines.Sparkline[];
                }

                export class SparklineSetting{
                    /**
                     * Creates the sparkline settings.
                     * @param {object} setting The settings.
                     * @class
                     */
                    constructor(setting?:  ISparklineSetting);
                    /**
                     * Indicates the options for the sparkline.<br />
                     * options.axisColor {string} the color of the axis<br />
                     * options.firstMarkerColor {string} the color of the first data point for each sparkline in this sparkline group<br />
                     * options.highMarkerColor {string} the color of the highest data point for each sparkline in this sparkline group<br />
                     * options.lastMarkerColor {string} the color of the last data point for each sparkline in this sparkline group<br />
                     * options.lowMarkerColor {string} the color of the lowest data point for each sparkline in this sparkline group<br />
                     * options.markersColor {string} a value that specifies the color of the data markers for each sparkline in this sparkline group<br />
                     * options.negativeColor {string} a value that specifies the color of the negative data points for each sparkline in this sparkline group<br />
                     * options.seriesColor {string} a value that specifies the color for each sparkline in this sparkline group<br />
                     * options.displayEmptyCellsAs {GC.Spread.Sheets.Sparklines.EmptyValueStyle} Indicates how to display the empty cells<br />
                     * options.rightToLeft {boolean} Indicates whether each sparkline in the sparkline group is displayed in a right-to-left manner<br />
                     * options.displayHidden {boolean} Indicates whether data in hidden cells is plotted for the sparklines in this sparkline group<br />
                     * options.displayXAxis {boolean} Indicates whether the horizontal axis is displayed for each sparkline in this sparkline group<br />
                     * options.showFirst {boolean} a value that indicates whether the first data point is formatted differently for each sparkline in this sparkline group<br />
                     * options.showHigh {boolean} a value that specifies whether the data points with the highest value are formatted differently for each sparkline in this sparkline group<br />
                     * options.showLast {boolean} a value that indicates whether the last data point is formatted differently for each sparkline in this sparkline group<br />
                     * options.showLow {boolean} a value that specifies whether the data points with the lowest value are formatted differently for each sparkline in this sparkline group<br />
                     * options.showNegative {boolean} a value that specifies whether the negative data points are formatted differently for each sparkline in this sparkline group<br />
                     * options.showMarkers {boolean} a value that specifies whether data markers are displayed for each sparkline in this sparkline group<br />
                     * options.manualMax {number} Indicates the maximum for the vertical axis that is shared across all sparklines in this sparkline group. The axis is zero if maxAxisType does not equal custom<br />
                     * options.manualMin {number} Indicates the minimum for the vertical axis that is shared across all sparklines in this sparkline group. The axis is zero if minAxisType does not equal custom<br />
                     * options.maxAxisType {GC.Spread.Sheets.Sparklines.SparklineAxisMinMax} Indicates how the vertical axis maximum is calculated for the sparklines in this sparkline group<br />
                     * options.minAxisType {GC.Spread.Sheets.Sparklines.SparklineAxisMinMax} Indicates how the vertical axis minimum is calculated for the sparklines in this sparkline group<br />
                     * options.groupMaxValue {number} Gets the maximum value of the sparkline group<br />
                     * options.groupMinValue {number} Gets the minimum value of the sparkline group<br />
                     * options.lineWeight {number} Indicates the line weight for each sparkline in the sparkline group, where the line weight is measured in points. The weight must be greater than or equal to zero, and must be less than or equal to 3 (LineSeries only supports line weight values in the range of 0.0-3.0)
                     */
                    options: Object;
                    /**
                     * Clones sparkline settings.
                     * @returns {GC.Spread.Sheets.Sparklines.SparklineSetting} The cloned sparkline setting.
                     */
                    clone(): GC.Spread.Sheets.Sparklines.SparklineSetting;
                }

                export class SpreadSparkline extends SparklineEx{
                    /**
                     * Represents the class for the Spread sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class StackedSparkline extends SparklineEx{
                    /**
                     * Represents the class for the stacked sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class VariSparkline extends SparklineEx{
                    /**
                     * Represents the class for the variance sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class VBarSparkline extends SparklineEx{
                    /**
                     * Represents the class for the Vbar sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class WinlossSparkline extends SparklineEx{
                    /**
                     * Represents the class for the winloss sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }

                export class YearSparkline extends SparklineEx{
                    /**
                     * Represents the class for the year sparkline.
                     * @extends GC.Spread.Sheets.Sparklines.SparklineEx
                     * @class
                     */
                    constructor();
                }
            }

            module Tables{
                /**
                 * Specifies what data is kept when removing the table.
                 * @enum {number}
                 */
                export enum TableRemoveOptions{
                    /**
                     *  Removes data and styles.
                     */
                    none= 0,
                    /**
                     *  Keeps values.
                     */
                    keepData= 1,
                    /**
                     *  Keeps styles.
                     */
                    keepStyle= 2
                }


                export class Table{
                    /**
                     * Represents a table that can be added in a sheet.
                     * @class
                     * @param {string} name The table name.
                     * @param {number} row The table row index.
                     * @param {number} col The table column index.
                     * @param {number} rowCount The table row count.
                     * @param {number} colCount The table column count.
                     * @param {GC.Spread.Sheets.Tables.TableTheme} style The table style.
                     * @param {Object} options The initialization options of the table.<br />
                     ** options.showHeader boolean Whether to display a header.<br />
                     ** options.showfooter boolean Whether to display a footer.
                     */
                    constructor(name?:  string,  row?:  number,  col?:  number,  rowCount?:  number,  colCount?:  number,  style?:  GC.Spread.Sheets.Tables.TableTheme);
                    /**
                     * Gets or sets whether to generate columns automatically while binding to a data source.
                     * @param {boolean} value Whether to generate columns automatically while binding to a data source.
                     * @returns {boolean | GC.Spread.Sheets.Tables.Table} If no value is set, returns whether to generate columns automatically while binding to a data source; otherwise, returns the table.
                     */
                    autoGenerateColumns(value?:  boolean): any;
                    /**
                     * Gets or sets a value that indicates whether to display an alternating column style.
                     * @param {boolean} value Whether to display an alternating column style.
                     * @returns {boolean | GC.Spread.Sheets.Tables.Table} If no value is set, returns whether to display an alternating column style; otherwise, returns the table.
                     */
                    bandColumns(value?:  boolean): any;
                    /**
                     * Gets or sets a value that indicates whether to display an alternating row style.
                     * @param {boolean} value Whether to display an alternating row style.
                     * @returns {boolean | GC.Spread.Sheets.Tables.Table} If no value is set, returns whether to display an alternating row style; otherwise, returns the table.
                     */
                    bandRows(value?:  boolean): any;
                    /**
                     * Binds the columns using the specified data fields.
                     * @param {Array} columns The array of table column information with data fields and names. Each item is GC.Spread.Sheets.Tables.TableColumn.
                     */
                    bindColumns(columns:  any[]): void;
                    /**
                     * Gets or sets the binding path for cell-level binding in the table.
                     * @param {string} value The binding path for cell-level binding in the table.
                     * @returns {string | GC.Spread.Sheets.Tables.Table} If no value is set, returns the binding path for cell-level binding in the table; otherwise, returns the table.
                     */
                    bindingPath(value?:  string): any;
                    /**
                     * Gets the cell range for the table data area.
                     * @returns {GC.Spread.Sheets.Range} The table data range.
                     */
                    dataRange(): GC.Spread.Sheets.Range;
                    /**
                     * Gets or sets whether the table column's filter button is displayed.
                     * @param {number} tableColumnIndex The table column index of the filter button.
                     * @param {boolean} value Whether the table column's filter button is displayed.
                     * @returns {boolean} The table column's filter button display state.
                     * @returns {boolean | GC.Spread.Sheets.Tables.Table}<br />
                     *  If no parameter is set, returns <c>false</c> if all filter buttons are invisible, otherwise, <c>true</c>.<br />
                     *  If only a number is set, returns whether the specified table column' filter button is displayed.<br />
                     *  If only a boolean that indicates whether to display filter buttons is set, applies to all filter buttons and returns the table.<br />
                     *  If two parameters are provided, applies to the specified table columns' filter button and returns the table.
                     */
                    filterButtonVisible(tableColumnIndex?:  number,  value?:  boolean): any;
                    /**
                     * Gets the footer index in the sheet.
                     * @returns {number} The footer index.
                     */
                    footerIndex(): number;
                    /**
                     * Gets the table footer formula with the specified index.
                     * @param {number} tableColumnIndex The column index of the table footer. The index is zero-based.
                     * @returns {string} The table footer formula.
                     */
                    getColumnFormula(tableColumnIndex:  number): string;
                    /**
                     * Gets the table header text with the specified table index.
                     * @param {number} tableColumnIndex The column index of the table header. The index is zero-based.
                     * @returns {string} The header text of the specified column by index.
                     */
                    getColumnName(tableColumnIndex:  number): any;
                    /**
                     * Gets the table footer value with the specified index.
                     * @param {number} tableColumnIndex The column index of the table footer. The index is zero-based.
                     * @returns {string} The table footer value.
                     */
                    getColumnValue(tableColumnIndex:  number): string;
                    /**
                     * Gets the header index in the sheet.
                     * @returns {number} The header index.
                     */
                    headerIndex(): number;
                    /**
                     * Gets or sets a value that indicates whether to highlight the first column.
                     * @param {boolean} value Whether to highlight the first column.
                     * @returns {boolean | GC.Spread.Sheets.Tables.Table} If no value is set, returns whether to highlight the first column; otherwise, returns the table.
                     */
                    highlightFirstColumn(value?:  boolean): any;
                    /**
                     * Gets or sets a value that indicates whether to highlight the last column.
                     * @param {boolean} value Whether to highlight the last column.
                     * @returns {boolean | GC.Spread.Sheets.Tables.Table} If no value is set, returns whether to highlight the last column; otherwise, returns the table.
                     */
                    highlightLastColumn(value?:  boolean): any;
                    /**
                     * Gets or sets the table name.
                     * @param {string} value The table name.
                     * @returns {string | GC.Spread.Sheets.Tables.Table} If no value is set, returns the table name; otherwise, returns the table.
                     */
                    name(value?:  string): any;
                    /**
                     * Gets the range for the entire table.
                     * @returns {GC.Spread.Sheets.Range} The whole table range.
                     */
                    range(): GC.Spread.Sheets.Range;
                    /**
                     * Gets the row filter for the table.
                     * @returns {GC.Spread.Sheets.Filter.HideRowFilter} The row filter.
                     */
                    rowFilter(): GC.Spread.Sheets.Filter.HideRowFilter;
                    /**
                     * Sets a formula to the table's data range with the specified index.
                     * @param {number} tableColumnIndex The column index of the table. The index is zero-based.
                     * @param {string} formula The data range formula.
                     * @returns {GC.Spread.Sheets.Tables.Table} The table.
                     */
                    setColumnDataFormula(tableColumnIndex:  number,  formula:  string): GC.Spread.Sheets.Tables.Table;
                    /**
                     * Sets the table footer formula with the specified index.
                     * @param {number} tableColumnIndex The column index of the table footer. The index is zero-based.
                     * @param {string} formula The table footer formula.
                     * @returns {GC.Spread.Sheets.Tables.Table} The table.
                     */
                    setColumnFormula(tableColumnIndex:  number,  formula:  string): GC.Spread.Sheets.Tables.Table;
                    /**
                     * Sets the table header text with the specified table index.
                     * @param {string} tableColumnIndex The column index of the table header. The index is zero-based.
                     * @param {string} name The header text.
                     * @returns {GC.Spread.Sheets.Tables.Table} The table.
                     */
                    setColumnName(tableColumnIndex:  number,  name:  string): GC.Spread.Sheets.Tables.Table;
                    /**
                     * Sets the table footer value with the specified index.
                     * @param {number} tableColumnIndex The column index of the table footer. The index is zero-based.
                     * @param {Object} value The table footer value.
                     * @returns {GC.Spread.Sheets.Tables.Table} The table.
                     */
                    setColumnValue(tableColumnIndex:  number,  value:  Object): GC.Spread.Sheets.Tables.Table;
                    /**
                     * Gets or sets a value that indicates whether to display a footer.
                     * @param {boolean} value Whether to display a footer.
                     * @returns {boolean | GC.Spread.Sheets.Tables.Table} If no value is set, returns whether to display a footer; otherwise, returns the table.
                     */
                    showFooter(value?:  boolean): any;
                    /**
                     * Gets or sets a value that indicates whether to display a header.
                     * @param {boolean} value Whether to display a header.
                     * @returns {boolean | GC.Spread.Sheets.Tables.Table} If no value is set, returns whether to display a header; otherwise, returns the table.
                     */
                    showHeader(value?:  boolean): any;
                    /**
                     * Gets or sets a style for the table.
                     * @param {GC.Spread.Sheets.Tables.TableTheme} value The style for the table.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme | GC.Spread.Sheets.Tables.Table} If no value is set, returns the table style; otherwise, returns the table.
                     */
                    style(value?:  GC.Spread.Sheets.Tables.TableTheme): any;
                }

                export class TableColumn{
                    /**
                     * Represents the table column information.
                     * @class
                     * @param {string} id The table column ID.
                     */
                    constructor(id:  string);
                    /**
                     * Gets or sets the table column data field for accessing the table's data source.
                     * @param {string} value The table column data field.
                     * @returns {string | GC.Spread.Sheets.Tables.TableColumn} If no value is set, returns the table column data field; otherwise, returns the table column.
                     */
                    dataField(value?:  string): any;
                    /**
                     * Gets or sets the table column ID.
                     * @param {number} value The table column ID.
                     * @returns {number | GC.Spread.Sheets.Tables.TableColumn} If no value is set, returns the table column ID; otherwise, returns the table column.
                     */
                    id(value?:  number): any;
                    /**
                     * Gets or sets the table column name for display.
                     * @param {string} value The table column name.
                     * @returns {string | GC.Spread.Sheets.Tables.TableColumn} If no value is set, returns the table column name; otherwise, returns the table column.
                     */
                    name(value?:  string): any;
                }

                export class TableManager{
                    /**
                     * Represents a table manager that can manage all tables in a sheet.
                     * @class
                     * @param {GC.Spread.Sheets.Worksheet} sheet The worksheet.
                     */
                    constructor(sheet:  GC.Spread.Sheets.Worksheet);
                    /**
                     * Adds a range table with a specified size to the sheet.
                     * @param {string} name The table name.
                     * @param {number} row The row index.
                     * @param {number} column The column index.
                     * @param {number} rowCount The row count of the table.
                     * @param {number} columnCount The column count of the table.
                     * @param {GC.Spread.Sheets.Tables.TableTheme} style The style of the table.
                     * @param {Object} options The initialization options of the table.<br />
                     ** options.showHeader boolean Whether to display a header.<br />
                     ** options.showfooter boolean Whether to display a footer.
                     * @returns {GC.Spread.Sheets.Tables.Table} The new table instance.
                     */
                    add(name?:  string,  row?:  number,  column?:  number,  rowCount?:  number,  columnCount?:  number,  style?:  GC.Spread.Sheets.Tables.TableTheme): GC.Spread.Sheets.Tables.Table;
                    /**
                     * Adds a range table with a specified data source to the sheet.
                     * @param {string} name The table name.
                     * @param {number} row The row index.
                     * @param {number} column The column index.
                     * @param {object} dataSource The data source for the table.
                     * @param {GC.Spread.Sheets.Tables.TableTheme} style The style of the table.
                     * @param {Object} options The initialization options of the table.<br />
                     ** options.showHeader boolean Whether to display a header.<br />
                     ** options.showfooter boolean Whether to display a footer.
                     * @returns {GC.Spread.Sheets.Tables.Table} The new table instance.
                     */
                    addFromDataSource(name:  string,  row:  number,  column:  number,  dataSource:  Object,  style:  GC.Spread.Sheets.Tables.TableTheme): GC.Spread.Sheets.Tables.Table;
                    /**
                     * Gets all tables of the sheet.
                     * @returns {Array} The GC.Spread.Sheets.Tables.Table array of table instances. The array is never null.
                     */
                    all(): GC.Spread.Sheets.Tables.Table[];
                    /**
                     * Gets the table of the specified cell.
                     * @param {number} row The row index.
                     * @param {number} column The column index.
                     * @returns {GC.Spread.Sheets.Tables.Table} The table instance if the cell belongs to a table; otherwise, <c>null</c>.
                     */
                    find(row:  number,  column:  number): GC.Spread.Sheets.Tables.Table;
                    /**
                     * Gets the table with a specified name.
                     * @param {string} name The table name.
                     * @returns {GC.Spread.Sheets.Tables.Table} The table instance if the cell belongs to a table; otherwise, <c>null</c>.
                     */
                    findByName(name:  string): GC.Spread.Sheets.Tables.Table;
                    /**
                     * Changes the table location.
                     * @param {GC.Spread.Sheets.Tables.Table|string} table The table instance or the table name.
                     * @param {number} row The new row index.
                     * @param {number} column The new column index.
                     */
                    move(table:  GC.Spread.Sheets.Tables.Table,  row:  number,  column:  number): void;
                    /**
                     * Removes a specified table.
                     * @param {GC.Spread.Sheets.Tables.Table|string} table The table instance or the table name.
                     * @param {GC.Spread.Sheets.Tables.TableRemoveOptions} options Specifies what data is kept when removing the table.
                     */
                    remove(table:  GC.Spread.Sheets.Tables.Table,  options:  GC.Spread.Sheets.Tables.TableRemoveOptions): void;
                    /**
                     * Changes the table size.
                     * @param {GC.Spread.Sheets.Tables.Table|string} table The table or the table name.
                     * @param {GC.Spread.Sheets.Range} range The new table range. The headers must remain in the same row, and the resulting table range must overlap the original table range.
                     */
                    resize(table:  GC.Spread.Sheets.Tables.Table,  range:  GC.Spread.Sheets.Range): void;
                }

                export class TableStyle{
                    /**
                     * Represents table style information.
                     * @class
                     * @param {string} backColor The background color of the table.
                     * @param {string} foreColor The foreground color of the table.
                     * @param {string} font The font.
                     * @param {GC.Spread.Sheets.LineBorder} borderLeft The left border line of the table.
                     * @param {GC.Spread.Sheets.LineBorder} borderTop The top border line of the table.
                     * @param {GC.Spread.Sheets.LineBorder} borderRight The right border line of the table.
                     * @param {GC.Spread.Sheets.LineBorder} borderBottom The bottom border line of the table.
                     * @param {GC.Spread.Sheets.LineBorder} borderHorizontal The horizontal border line of the table.
                     * @param {GC.Spread.Sheets.LineBorder} borderVertical The vertical border line of the table.
                     * @param {GC.Spread.Sheets.TextDecorationType} textDecoration The text decoration of the table.
                     */
                    constructor(backColor?:  string,  foreColor?:  string,  font?:  string,  borderLeft?:  GC.Spread.Sheets.LineBorder,  borderTop?:  GC.Spread.Sheets.LineBorder,  borderRight?:  GC.Spread.Sheets.LineBorder,  borderBottom?:  GC.Spread.Sheets.LineBorder,  borderHorizontal?:  GC.Spread.Sheets.LineBorder,  borderVertical?:  GC.Spread.Sheets.LineBorder,  textDecoration?:  GC.Spread.Sheets.TextDecorationType);
                    /**
                     * Indicates the background color.
                     * @type {string}
                     */
                    backColor: string;
                    /**
                     * Indicates the bottom border line of the table.
                     * @type {GC.Spread.Sheets.LineBorder}
                     */
                    borderBottom: GC.Spread.Sheets.LineBorder;
                    /**
                     * Indicates the horizontal border line of the table.
                     * @type {GC.Spread.Sheets.LineBorder}
                     */
                    borderHorizontal: GC.Spread.Sheets.LineBorder;
                    /**
                     * Indicates the left border line of the table.
                     * @type {GC.Spread.Sheets.LineBorder}
                     */
                    borderLeft: GC.Spread.Sheets.LineBorder;
                    /**
                     * Indicates the right border line of the table.
                     * @type {GC.Spread.Sheets.LineBorder}
                     */
                    borderRight: GC.Spread.Sheets.LineBorder;
                    /**
                     * Indicates the top border line of the table.
                     * @type {GC.Spread.Sheets.LineBorder}
                     */
                    borderTop: GC.Spread.Sheets.LineBorder;
                    /**
                     * Indicates the vertical border line of the table.
                     * @type {GC.Spread.Sheets.LineBorder}
                     */
                    borderVertical: GC.Spread.Sheets.LineBorder;
                    /**
                     * Indicates the font.
                     * @type {string}
                     */
                    font: string;
                    /**
                     * Indicates the foreground color.
                     * @type {string}
                     */
                    foreColor: string;
                    /**
                     * Indicates the text decoration of the table.
                     * @type {GC.Spread.Sheets.TextDecorationType}
                     */
                    textDecoration: GC.Spread.Sheets.TextDecorationType;
                }

                export class TableTheme{
                    /**
                     * Represents the table style settings.
                     * @class
                     */
                    constructor();
                    /**
                     * Gets or sets the size of the first alternating column.
                     * @param {number} value The size of the first alternating column.
                     * @returns {number | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the size of the first alternating column; otherwise, returns the table theme.
                     */
                    firstColumnStripSize(value?:  number): any;
                    /**
                     * Gets or sets the style of the first alternating column.
                     * @param {GC.Spread.Sheets.Tables.TableStyle} value The style of the first alternating column.
                     * @returns {GC.Spread.Sheets.Tables.TableStyle | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the style of the first alternating column; otherwise, returns the table theme.
                     */
                    firstColumnStripStyle(value?:  GC.Spread.Sheets.Tables.TableStyle): any;
                    /**
                     * Gets or sets the style of the first footer cell.
                     * @param {GC.Spread.Sheets.Tables.TableStyle} value The style of the first footer cell.
                     * @returns {GC.Spread.Sheets.Tables.TableStyle | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the style of the first footer cell; otherwise, returns the table theme.
                     */
                    firstFooterCellStyle(value?:  GC.Spread.Sheets.Tables.TableStyle): any;
                    /**
                     * Gets or sets the style of the first header cell.
                     * @param {GC.Spread.Sheets.Tables.TableStyle} value The style of the first header cell.
                     * @returns {GC.Spread.Sheets.Tables.TableStyle | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the style of the first header cell; otherwise, returns the table theme.
                     */
                    firstHeaderCellStyle(value?:  GC.Spread.Sheets.Tables.TableStyle): any;
                    /**
                     * Gets or sets the size of the first alternating row.
                     * @param {number} value The size of the first alternating row.
                     * @returns {number | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the size of the first alternating row; otherwise, returns the table theme.
                     */
                    firstRowStripSize(value?:  number): any;
                    /**
                     * Gets or sets the first alternating row style.
                     * @param {GC.Spread.Sheets.Tables.TableStyle} value The first alternating row style.
                     * @returns {GC.Spread.Sheets.Tables.TableStyle | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the first alternating row style; otherwise, returns the table theme.
                     */
                    firstRowStripStyle(value?:  GC.Spread.Sheets.Tables.TableStyle): any;
                    /**
                     * Gets or sets the default style of the footer area.
                     * @param {GC.Spread.Sheets.Tables.TableStyle} value The default style of the footer area.
                     * @returns {GC.Spread.Sheets.Tables.TableStyle | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the default style of the footer area; otherwise, returns the table theme.
                     */
                    footerRowStyle(value?:  GC.Spread.Sheets.Tables.TableStyle): any;
                    /**
                     * Gets or sets the default style of the header area.
                     * @param {GC.Spread.Sheets.Tables.TableStyle} value The default style of the header area.
                     * @returns {GC.Spread.Sheets.Tables.TableStyle | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the default style of the header area; otherwise, returns the table theme.
                     */
                    headerRowStyle(value?:  GC.Spread.Sheets.Tables.TableStyle): any;
                    /**
                     * Gets or sets the style of the first column.
                     * @param {GC.Spread.Sheets.Tables.TableStyle} value The style of the first column.
                     * @returns {GC.Spread.Sheets.Tables.TableStyle | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the style of the first column; otherwise, returns the table theme.
                     */
                    highlightFirstColumnStyle(value?:  GC.Spread.Sheets.Tables.TableStyle): any;
                    /**
                     * Gets or sets the style of the last column.
                     * @param {GC.Spread.Sheets.Tables.TableStyle} value The style of the last column.
                     * @returns {GC.Spread.Sheets.Tables.TableStyle | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the style of the last column; otherwise, returns the table theme.
                     */
                    highlightLastColumnStyle(value?:  GC.Spread.Sheets.Tables.TableStyle): any;
                    /**
                     * Gets or sets the style of the last footer cell.
                     * @param {GC.Spread.Sheets.Tables.TableStyle} value The style of the last footer cell.
                     * @returns {GC.Spread.Sheets.Tables.TableStyle | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the style of the last footer cell; otherwise, returns the table theme.
                     */
                    lastFooterCellStyle(value?:  GC.Spread.Sheets.Tables.TableStyle): any;
                    /**
                     * Gets or sets the style of the last header cell.
                     * @param {GC.Spread.Sheets.Tables.TableStyle} value The style of the last header cell.
                     * @returns {GC.Spread.Sheets.Tables.TableStyle | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the style of the last header cell; otherwise, returns the table theme.
                     */
                    lastHeaderCellStyle(value?:  GC.Spread.Sheets.Tables.TableStyle): any;
                    /**
                     * Gets or sets the name of the style.
                     * @param {string} value The name of the style.
                     * @returns {string | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the name of the style; otherwise, returns the table theme.
                     */
                    name(value?:  string): any;
                    /**
                     * Gets or sets the size of the second alternating column.
                     * @param {number} value The size of the second alternating column.
                     * @returns {number | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the size of the second alternating column; otherwise, returns the table theme.
                     */
                    secondColumnStripSize(value?:  number): any;
                    /**
                     * Gets or sets the style of the second alternating column.
                     * @param {GC.Spread.Sheets.Tables.TableStyle} value The style of the second alternating column.
                     * @returns {GC.Spread.Sheets.Tables.TableStyle | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the style of the second alternating column; otherwise, returns the table theme.
                     */
                    secondColumnStripStyle(value?:  GC.Spread.Sheets.Tables.TableStyle): any;
                    /**
                     * Gets or sets the size of the second alternating row.
                     * @param {number} value The size of the second alternating row.
                     * @returns {number | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the size of the second alternating row; otherwise, returns the table theme.
                     */
                    secondRowStripSize(value?:  number): any;
                    /**
                     * Gets or sets the second alternating row style.
                     * @param {GC.Spread.Sheets.Tables.TableStyle} value The second alternating row style.
                     * @returns {GC.Spread.Sheets.Tables.TableStyle | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the second alternating row style; otherwise, returns the table theme.
                     */
                    secondRowStripStyle(value?:  GC.Spread.Sheets.Tables.TableStyle): any;
                    /**
                     * Gets or sets the default style of the data area.
                     * @param {GC.Spread.Sheets.Tables.TableStyle} value The default style of the data area.
                     * @returns {GC.Spread.Sheets.Tables.TableStyle | GC.Spread.Sheets.Tables.TableTheme} If no value is set, returns the default style of the data area; otherwise, returns the table theme.
                     */
                    wholeTableStyle(value?:  GC.Spread.Sheets.Tables.TableStyle): any;
                }

                export class TableThemes{
                    /**
                     * Represents a built-in table theme collection.
                     * @class
                     */
                    constructor();
                    /**
                     * Gets the dark1 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static dark1: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the dark10 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static dark10: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the dark11 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static dark11: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the dark2 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static dark2: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the dark3 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static dark3: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the dark4 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static dark4: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the dark5 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static dark5: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the dark6 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static dark6: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the dark7 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static dark7: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the dark8 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static dark8: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the dark9 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static dark9: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light1 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light1: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light10 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light10: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light11 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light11: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light12 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light12: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light13 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light13: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light14 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light14: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light15 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light15: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light16 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light16: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light17 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light17: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light18 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light18: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light19 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light19: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light2 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light2: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light20 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light20: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light21 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light21: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light3 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light3: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light4 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light4: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light5 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light5: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light6 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light6: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light7 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light7: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light8 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light8: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the light9 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static light9: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium1 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium1: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium10 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium10: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium11 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium11: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium12 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium12: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium13 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium13: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium14 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium14: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium15 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium15: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium16 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium16: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium17 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium17: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium18 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium18: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium19 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium19: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium2 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium2: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium20 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium20: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium21 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium21: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium22 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium22: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium23 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium23: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium24 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium24: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium25 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium25: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium26 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium26: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium27 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium27: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium28 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium28: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium3 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium3: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium4 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium4: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium5 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium5: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium6 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium6: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium7 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium7: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium8 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium8: GC.Spread.Sheets.Tables.TableTheme;
                    /**
                     * Gets the medium9 style.
                     * @returns {GC.Spread.Sheets.Tables.TableTheme}
                     */
                    static medium9: GC.Spread.Sheets.Tables.TableTheme;
                }
            }

            module Touch{

                export class TouchToolStrip{
                    /**
                     * Represents a toolbar.
                     * @class
                     * @param {GC.Spread.Sheets.Workbook} workbook The Spread object.
                     * @param {HTMLElement} host The host DOM element.
                     */
                    constructor(workbook:  GC.Spread.Sheets.Workbook,  host:  HTMLElement);
                    /**
                     * Adds an item to the touch toolbar.
                     * @param {any} item The item to be added.
                     * @remarks The item to be added can be a toolbar item or a line separator.
                     */
                    add(item:  any): void;
                    /**
                     * Clears all items in the toolbar.
                     */
                    clear(): void;
                    /**
                     * Closes the toolbar.
                     */
                    close(): void;
                    /**
                     * Gets the item with the specified name.
                     * @param {string} name The item name.
                     * @returns If the item exists in the toolbar, the item is returned; otherwise, returns undefined.
                     */
                    getItem(name:  string): any;
                    /**
                     * Gets all the items that belong to the toolbar.
                     * @returns An array that contains all the items in the toolbar.
                     */
                    getItems(): any;
                    /**
                     * Gets or sets the image area height.
                     * @param {number} height The image area height.
                     * @returns {number | GC.Spread.Sheets.Touch.TouchToolStrip} If no value is set, returns the image area height; otherwise, returns the toolbar.
                     */
                    imageAreaHeight(height?:  number): any;
                    /**
                     * Gets or sets the toolbar item height.
                     * @param {number} height The toolbar item height.
                     * @returns {number | GC.Spread.Sheets.Touch.TouchToolStrip} If no value is set, returns the toolbar item height; otherwise, returns the toolbar.
                     */
                    itemHeight(height?:  number): any;
                    /**
                     * Gets or sets the toolbar item width.
                     * @param {number} width The toolbar item width.
                     * @returns {number | GC.Spread.Sheets.Touch.TouchToolStrip} If no value is set, returns the toolbar item width; otherwise, returns the toolbar.
                     */
                    itemWidth(width?:  number): any;
                    /**
                     * Opens a toolbar in a specific position relative to the touch point.
                     * @param {number} x The <i>x</i>-coordinate.
                     * @param {number} y The <i>y</i>-coordinate.
                     */
                    open(x:  number,  y:  number): void;
                    /**
                     * Removes the toolbar item with the specified name.
                     * @param {string} name The name of the item to be removed.
                     * @returns {GC.Spread.Sheets.Touch.TouchToolStripItem} The removed item.
                     */
                    remove(name:  string): GC.Spread.Sheets.Touch.TouchToolStripItem;
                    /**
                     * Gets or sets the toolbar separator height.
                     * @param {number} height The toolbar separator height.
                     * @returns {number | GC.Spread.Sheets.Touch.TouchToolStrip} If no value is set, returns the toolbar separator height; otherwise, returns the toolbar.
                     */
                    separatorHeight(height?:  number): any;
                }

                export class TouchToolStripItem{
                    /**
                     * Represents an item in the toolbar.
                     * @class
                     * @param {string} name The name of the item.
                     * @param {string} text The item text.
                     * @param {string} image The item image source.
                     * @param {any} command Defines the executive function that occurs when the user taps the item.
                     * @param {any} canExecute Defines when to show the item by a function. If returns <c>true</c>, display the item; otherwise, hide the item.
                     */
                    constructor(name:  string,  text:  string,  image:  string,  command?:  any,  canExecute?:  any);
                    /**
                     * Gets or sets the font of the item text.
                     * @param {string} value The font of the toolbar item text.
                     * @returns {string | GC.Spread.Sheets.Touch.TouchToolStripItem} If no value is set, returns the font of the item text; otherwise, returns the toolbar item.
                     */
                    font(value?:  string): any;
                    /**
                     * Gets or sets the color of the item text.
                     * @param {string} value The color of the toolbar item text.
                     * @returns {string | GC.Spread.Sheets.Touch.TouchToolStripItem} If no value is set, returns the color of the item text; otherwise, returns the toolbar item.
                     */
                    foreColor(value?:  string): any;
                    /**
                     * Gets or sets the source of the item image.
                     * @param {string} value The path and filename for the item image source.
                     * @returns {string | GC.Spread.Sheets.Touch.TouchToolStripItem} If no value is set, returns the source of the item image; otherwise, returns the toolbar item.
                     */
                    image(value?:  string): any;
                    /**
                     * Gets or sets the name of the item.
                     * @param {string} value The name of the toolbar item.
                     * @returns {string | GC.Spread.Sheets.Touch.TouchToolStripItem} If no value is set, returns the name of the item; otherwise, returns the toolbar item.
                     */
                    name(value?:  string): any;
                    /**
                     * Gets or sets the text of the item.
                     * @param {string} value The text of the toolbar item.
                     * @returns {string | GC.Spread.Sheets.Touch.TouchToolStripItem} If no value is set, returns the text of the item; otherwise, returns the toolbar item.
                     */
                    text(value?:  string): any;
                }

                export class TouchToolStripSeparator{
                    /**
                     * Represents a separator in the toolbar.
                     * @class
                     * @param {string} canExecute Defines when to display the separator with a function. If returns <c>true</c>, display the separator; otherwise, hide the separator.
                     */
                    constructor(canExecute?:  string);
                    /**
                     * Gets the name of the separator.
                     * @returns {string} Returns the current separator name.
                     */
                    name(): string;
                }
            }

        }

        module Slicers{

            export interface ISlicerConditional{
                exclusiveRowIndexes?: number[];
                ranges?: ISlicerRangeConditional[];
            }


            export interface ISlicerData{
                getColumnIndex(columnName: string): number;
                getData(columnName: string, range?: ISlicerRangeConditional): any[];
                getExclusiveData(columnName: string): any[];
                getRowIndexes(columnName: string, exclusiveRowIndex: number): number[];
                getExclusiveRowIndex(columnName: string, rowIndex: number): number;
                getFilteredIndexes(columnName: string, isPreview?: boolean): number[];
                getFilteredOutIndexes(columnName: string, filteredOutDataType: GC.Spread.Slicers.FilteredOutDataType, isPreview?: boolean): number[];
                getFilteredRanges(columnName: string): ISlicerRangeConditional[];
                getFilteredOutRanges(columnName: string): ISlicerRangeConditional[];
                attachListener(listener: ISlicerListener): void;
                detachListener(listener: ISlicerListener): void;
                doFilter(columnName: string, slicerConditional: ISlicerConditional, isPreview?: boolean): void;
                doUnfilter(columnName: string): void;
                clearPreview():void;
            }


            export interface ISlicerDataItem{
                columnName: string;
                rowIndex: number;
                data: any;
            }


            export interface ISlicerFiltedData{
                isPreview: boolean;
                rowIndexes: number[];
            }


            export interface ISlicerListener{
                onFiltered(data: ISlicerFiltedData): void;
                onDataChanged(data: ISlicerDataItem[]): void;
                onRowsChanged(rowIndex: number, rowCount: number, isAdd: boolean): void;
                onColumnNameChanged(oldName: string, newName: string): void;
                onColumnRemoved(columnName: string): void;
            }


            export interface ISlicerRangeConditional{
                min: number;
                max: number;
            }

            /**
             * Represents the kind of filtered out exclusive data index that should be included in the result.
             * @enum {number}
             */
            export enum FilteredOutDataType{
                /**
                 * Indicates all of the filtered out data.
                 */
                all= 0,
                /**
                 * Indicates the data was filtered out based on the current column.
                 */
                byCurrentColumn= 1,
                /**
                 * Indicates the data was filtered out based on other columns.
                 */
                byOtherColumns= 2
            }

            /**
             * Represents the aggregate type.
             * @enum {number}
             */
            export enum SlicerAggregateType{
                /**
                 *  Calculates the average of the specified numeric values.
                 */
                average= 1,
                /**
                 *  Calculates the number of data that contain numbers.
                 */
                count= 2,
                /**
                 *  Calculates the number of data that contain non-null values.
                 */
                counta= 3,
                /**
                 *  Calculates the maximum value, the greatest value, of all the values.
                 */
                max= 4,
                /**
                 *  Calculates the minimum value, the least value, of all the values.
                 */
                min= 5,
                /**
                 *  Multiplies all the arguments and returns the product.
                 */
                product= 6,
                /**
                 *  Calculates the standard deviation based on a sample.
                 */
                stdev= 7,
                /**
                 *  Calculates the standard deviation of a population based on the entire population using the numbers in a column of a list or database that match the specified conditions.
                 */
                stdevp= 8,
                /**
                 *  Calculates the sum of the specified numeric values.
                 */
                sum= 9,
                /**
                 *  Calculates the variance based on a sample of a population, which uses only numeric values.
                 */
                vars= 10,
                /**
                 *  Calculates the variance based on a sample of a population, which includes numeric, logical, or text values.
                 */
                varp= 11
            }


            export class GeneralSlicerData{
                /**
                 * Represents general slicer data.
                 * @class GC.Spread.Slicers.GeneralSlicerData
                 * @param {Array.<Array.<any>>} data The slicer data; it is a matrix array.
                 * @param {Array} columnNames The column names of the slicer data.
                 */
                constructor(data:  any[][],  columnNames:  string[]);
                /**
                 * Indicates the column names for the general slicer data.
                 * @type {Array}
                 */
                columnNames: string[];
                /**
                 * Indicates the data source for general slicer data.
                 * @type {{Array.<Array.<any>>}}
                 */
                data: any[][];
                /**
                 * Aggregates the data by the specified column name.
                 * @param {string} columnName The column name.
                 * @param {GC.Spread.Slicers.SlicerAggregateType} aggregateType The aggregate type.
                 * @param {object} range The specific range.<br />
                 * range.min: number type, the minimum value.<br />
                 * range.max: number type, the maximum value.
                 * @returns {number} The aggregated data.
                 */
                aggregateData(columnName:  string,  aggregateType:  GC.Spread.Slicers.SlicerAggregateType,  range?:  GC.Spread.Slicers.ISlicerRangeConditional): number;
                /**
                 * Attaches the listener.
                 * @param {GC.Spread.Slicers.ISlicerListener} listener The listener.
                 */
                attachListener(listener:  GC.Spread.Slicers.ISlicerListener): void;
                /**
                 * Clears the preview filter state.
                 */
                clearPreview(): any;
                /**
                 * Detaches the listener.
                 * @param {GC.Spread.Slicers.ISlicerListener} listener The listener.
                 */
                detachListener(listener:  GC.Spread.Slicers.ISlicerListener): void;
                /**
                 * Filters the data that corresponds to the specified column name and exclusive data indexes.
                 * @param {string} columnName The column name.
                 * @param {object} conditional The conditional filter.<br />
                 * conditional.exclusiveRowIndexes: number array type, visible exclusive row indexes<br />
                 * conditional.ranges: {min:number, max:number} array type, visible ranges.
                 * @param {boolean} isPreview Indicates whether preview is set.
                 */
                doFilter(columnName:  string,  conditional:  GC.Spread.Slicers.ISlicerConditional,  isPreview?:  boolean): void;
                /**
                 * Unfilters the data that corresponds to the specified column name.
                 * @param {string} columnName The column name.
                 */
                doUnfilter(columnName:  string): void;
                /**
                 * Gets the column index by the specified column name.
                 * @param {string} columnName The column name.
                 * @returns {number} The column index.
                 */
                getColumnIndex(columnName:  string): number;
                /**
                 * Gets the data by the specified column name.
                 * @param {string} columnName The column name.
                 * @param {object} range The specific range.
                 * range.min: number type, the minimum value.
                 * range.max: number type, the maximum value.
                 * @returns {Array} The data that corresponds to the specified column name.
                 */
                getData(columnName:  string,  range?:  GC.Spread.Slicers.ISlicerRangeConditional): any[];
                /**
                 * Gets the exclusive data by the specified column name.
                 * @param {string} columnName The column name.
                 * @returns {Array} The exclusive data that corresponds to the specified column name.
                 */
                getExclusiveData(columnName:  string): any[];
                /**
                 * Gets the exclusive data index by the specified column name and data index.
                 * @param {string} columnName The column name.
                 * @param {number} rowIndex The index of the data.
                 * @returns {number} The exclusive data index that corresponds to the specified column name and data index.
                 */
                getExclusiveRowIndex(columnName:  string,  rowIndex:  number): number;
                /**
                 * Gets the filtered exclusive data indexes by the specified column name.
                 * @param {string} columnName The column name.
                 * @returns {Array} The filtered exclusive data indexes that correspond to the specified column name.
                 */
                getFilteredIndexes(columnName:  string): number[];
                /**
                 * Gets the filtered out exclusive data indexes by the specified column name.
                 * @param {string} columnName The column name.
                 * @param {GC.Spread.Slicers.FilteredOutDataType} filteredOutDataType Indicates the kind of filtered out exclusive data index that should be included in the result.
                 * @returns {Array} The filtered out exclusive data indexes that correspond to the specified column name.
                 */
                getFilteredOutIndexes(columnName:  string,  filteredOutDataType:  FilteredOutDataType): number[];
                /**
                 * Gets the filtered out ranges by other columns.
                 * @param {string} columnName The column name.
                 * @returns {Array} The filtered out ranges by other columns that correspond to the specified column name.
                 */
                getFilteredOutRanges(columnName:  string): GC.Spread.Slicers.ISlicerRangeConditional[];
                /**
                 * Gets the filtered out row indexes.
                 * @returns {Array} The filtered out row indexes.
                 */
                getFilteredOutRowIndexes(): number[];
                /**
                 * Gets the filtered ranges by the specified column name.
                 * @param {string} columnName The column name.
                 * @returns {Array} The filtered ranges that correspond to the specified column name.
                 */
                getFilteredRanges(columnName:  string): GC.Spread.Slicers.ISlicerRangeConditional[];
                /**
                 * Gets the filtered row indexes.
                 * @returns {Array} The filtered row indexes.
                 */
                getFilteredRowIndexes(): number[];
                /**
                 * Gets the data indexes by the specified column name and exclusive data index.
                 * @param {string} columnName The column name.
                 * @param {number} exclusiveRowIndex The index of the exclusive data.
                 * @returns {Array} The data indexes that correspond to the specified column name and exclusive data index.
                 */
                getRowIndexes(columnName:  string,  exclusiveRowIndex:  number): number[];
                /**
                 * Gets whether the slicer is in the preview state.
                 */
                inPreview(): any;
                /**
                 * Changes a column name for the general slicer data.
                 * @param {string} oldName The old name of the column.
                 * @param {string} newName The new name of the column.
                 */
                onColumnNameChanged(oldName:  string,  newName:  string): void;
                /**
                 * Removes columns of the general slicer data.
                 * @param {number} colIndex The index of the starting column.
                 * @param {number} colCount The number of columns to remove.
                 */
                onColumnsRemoved(colIndex:  number,  colCount:  number): void;
                /**
                 * Changes data items in the data source of the general slicer data.
                 * @param {GC.Spread.Slicers.ISlicerDataItem} changedData The changed data item in the data source.
                 */
                onDataChanged(changedDataItems:  GC.Spread.Slicers.ISlicerDataItem): void;
                /**
                 * Occurs after the slicer data has been filtered.
                 * @param {Array} filteredIndexes The filtered exclusive data indexes.
                 * @param {boolean} isPreview Indicates whether the slicer is in preview mode.
                 */
                onFiltered(): any;
                /**
                 * Adds rows in the data source of the general slicer data.
                 * @param {number} rowIndex The index of the starting row.
                 * @param {number} rowCount The number of rows to add.
                 */
                onRowsAdded(rowIndex:  number,  rowCount:  number): void;
                /**
                 * Removes rows in the data source of the general slicer data.
                 * @param {number} rowIndex The index of the starting row.
                 * @param {number} rowCount The number of rows to remove.
                 */
                onRowsRemoved(rowIndex:  number,  rowCount:  number): void;
                /**
                 * Resumes the onFiltered event.
                 */
                resumeFilteredEvents(): any;
                /**
                 * Suspends the onFiltered event.
                 */
                suspendFilteredEvents(): any;
            }
        }

    }

}
