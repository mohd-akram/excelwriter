declare namespace XLSX {
  enum Alignment {
    /** No alignment. Cell will use Excel's default for the data type */
    NONE_ALIGN,
    /** Left horizontal alignment */
    LEFT_ALIGN,
    /** Center horizontal alignment */
    CENTER_ALIGN,
    /** Right horizontal alignment */
    RIGHT_ALIGN,
    /** Cell fill horizontal alignment */
    FILL_ALIGN,
    /** Justify horizontal alignment */
    JUSTIFY_ALIGN,
    /** Center Across horizontal alignment */
    CENTER_ACROSS_ALIGN,
    /** Left horizontal alignment */
    DISTRIBUTED_ALIGN,
    /** Top vertical alignment */
    VERTICAL_TOP_ALIGN,
    /** Bottom vertical alignment */
    VERTICAL_BOTTOM_ALIGN,
    /** Center vertical alignment */
    VERTICAL_CENTER_ALIGN,
    /** Justify vertical alignment */
    VERTICAL_JUSTIFY_ALIGN,
    /** Distributed vertical alignment */
    VERTICAL_DISTRIBUTED_ALIGN,
  }

  enum BorderStyle {
    /** No border */
    NONE_BORDER,
    /** Thin border style */
    THIN_BORDER,
    /** Medium border style */
    MEDIUM_BORDER,
    /** Dashed border style */
    DASHED_BORDER,
    /** Dotted border style */
    DOTTED_BORDER,
    /** Thick border style */
    THICK_BORDER,
    /** Double border style */
    DOUBLE_BORDER,
    /** Hair border style */
    HAIR_BORDER,
    /** Medium dashed border style */
    MEDIUM_DASHED_BORDER,
    /** Dash-dot border style */
    DASH_DOT_BORDER,
    /** Medium dash-dot border style */
    MEDIUM_DASH_DOT_BORDER,
    /** Dash-dot-dot border style */
    DASH_DOT_DOT_BORDER,
    /** Medium dash-dot-dot border style */
    MEDIUM_DASH_DOT_DOT_BORDER,
    /** Slant dash-dot border style */
    SLANT_DASH_DOT_BORDER,
  }

  enum ChartType {
    /** Area chart. */
    AREA_CHART = 1,
    /** Area chart - stacked. */
    AREA_STACKED_CHART,
    /** Area chart - percentage stacked. */
    AREA_STACKED_PERCENT_CHART,
    /** Bar chart. */
    BAR_CHART,
    /** Bar chart - stacked. */
    BAR_STACKED_CHART,
    /** Bar chart - percentage stacked. */
    BAR_STACKED_PERCENT_CHART,
    /** Column chart. */
    COLUMN_CHART,
    /** Column chart - stacked. */
    COLUMN_STACKED_CHART,
    /** Column chart - percentage stacked. */
    COLUMN_STACKED_PERCENT_CHART,
    /** Doughnut chart. */
    DOUGHNUT_CHART,
    /** Line chart. */
    LINE_CHART,
    /** Line chart - stacked. */
    LINE_STACKED_CHART,
    /** Line chart - percentage stacked. */
    LINE_STACKED_PERCENT_CHART,
    /** Pie chart. */
    PIE_CHART,
    /** Scatter chart. */
    SCATTER_CHART,
    /** Scatter chart - straight. */
    SCATTER_STRAIGHT_CHART,
    /** Scatter chart - straight with markers. */
    SCATTER_STRAIGHT_WITH_MARKERS_CHART,
    /** Scatter chart - smooth. */
    SCATTER_SMOOTH_CHART,
    /** Scatter chart - smooth with markers. */
    SCATTER_SMOOTH_WITH_MARKERS_CHART,
    /** Radar chart. */
    RADAR_CHART,
    /** Radar chart - with markers. */
    RADAR_WITH_MARKERS_CHART,
    /** Radar chart - filled. */
    RADAR_FILLED_CHART,
  }

  enum ScriptStyle {
    /** Superscript font */
    SUPERSCRIPT_FONT = 1,
    /** Subscript font */
    SUBSCRIPT_FONT,
  }

  enum UnderlineStyle {
    NONE_UNDERLINE,
    /** Single underline */
    SINGLE_UNDERLINE,
    /** Double underline */
    DOUBLE_UNDERLINE,
    /** Single accounting underline */
    SINGLE_ACCOUNTING_UNDERLINE,
    /** Double accounting underline */
    DOUBLE_ACCOUNTING_UNDERLINE,
  }

  enum FilterCriteria {
    /** Filter cells equal to a value. */
    EQUAL_TO_FILTER_CRITERIA = 1,
    /** Filter cells not equal to a value. */
    NOT_EQUAL_TO_FILTER_CRITERIA,
    /** Filter cells greater than a value. */
    GREATER_THAN_FILTER_CRITERIA,
    /** Filter cells less than a value. */
    LESS_THAN_FILTER_CRITERIA,
    /** Filter cells greater than or equal to a value. */
    GREATER_THAN_OR_EQUAL_TO_FILTER_CRITERIA,
    /** Filter cells less than or equal to a value. */
    LESS_THAN_OR_EQUAL_TO_FILTER_CRITERIA,
    /** Filter cells that are blank. */
    BLANKS_FILTER_CRITERIA,
    /** Filter cells that are not blank. */
    NON_BLANKS_FILTER_CRITERIA,
  }

  enum FilterOperator {
    /** Logical "and" of 2 filter rules. */
    AND_FILTER,
    /** Logical "or" of 2 filter rules. */
    OR_FILTER,
  }

  enum ValidationType {
    /** Restrict cell input to whole/integer numbers only. */
    INTEGER_VALIDATION_TYPE = 1,
    /** Restrict cell input to whole/integer numbers only, using a cell reference. */
    INTEGER_FORMULA_VALIDATION_TYPE,
    /** Restrict cell input to decimal numbers only. */
    DECIMAL_VALIDATION_TYPE,
    /** Restrict cell input to decimal numbers only, using a cell reference. */
    DECIMAL_FORMULA_VALIDATION_TYPE,
    /** Restrict cell input to a list of strings in a dropdown. */
    LIST_VALIDATION_TYPE,
    /** Restrict cell input to a list of strings in a dropdown, using a cell range. */
    LIST_FORMULA_VALIDATION_TYPE,
    /** Restrict cell input to date values only, using a lxw_datetime type. */
    DATE_VALIDATION_TYPE,
    /** Restrict cell input to date values only, using a cell reference. */
    DATE_FORMULA_VALIDATION_TYPE,
    /* Restrict cell input to date values only, as a serial number. Undocumented. */
    DATE_NUMBER_VALIDATION_TYPE,
    /** Restrict cell input to time values only, using a lxw_datetime type. */
    TIME_VALIDATION_TYPE,
    /** Restrict cell input to time values only, using a cell reference. */
    TIME_FORMULA_VALIDATION_TYPE,
    /* Restrict cell input to time values only, as a serial number. Undocumented. */
    TIME_NUMBER_VALIDATION_TYPE,
    /** Restrict cell input to strings of defined length. */
    LENGTH_VALIDATION_TYPE,
    /** Restrict cell input to strings of defined length, using a cell reference. */
    LENGTH_FORMULA_VALIDATION_TYPE,
    /** Restrict cell to input controlled by a custom formula that returns `TRUE/FALSE`. */
    CUSTOM_FORMULA_VALIDATION_TYPE,
    /** Allow any type of input. Mainly only useful for pop-up messages. */
    ANY_VALIDATION_TYPE,
  }

  enum ValidationCriteria {
    /** Select data between two values. */
    BETWEEN_VALIDATION_CRITERIA = 1,
    /** Select data that is not between two values. */
    NOT_BETWEEN_VALIDATION_CRITERIA,
    /** Select data equal to a value. */
    EQUAL_TO_VALIDATION_CRITERIA,
    /** Select data not equal to a value. */
    NOT_EQUAL_TO_VALIDATION_CRITERIA,
    /** Select data greater than a value. */
    GREATER_THAN_VALIDATION_CRITERIA,
    /** Select data less than a value. */
    LESS_THAN_VALIDATION_CRITERIA,
    /** Select data greater than or equal to a value. */
    GREATER_THAN_OR_EQUAL_TO_VALIDATION_CRITERIA,
    /** Select data less than or equal to a value. */
    LESS_THAN_OR_EQUAL_TO_VALIDATION_CRITERIA,
  }

  enum ValidationErrorType {
    /** Show a "Stop" data validation pop-up message. This is the default. */
    STOP_VALIDATION_ERROR_TYPE,
    /** Show an "Error" data validation pop-up message. */
    WARNING_VALIDATION_ERROR_TYPE,
    /** Show an "Information" data validation pop-up message. */
    INFORMATION_VALIDATION_ERROR_TYPE,
  }

  type Chart = ExcelWriter.Chart;
  type Format = ExcelWriter.Format;
}

declare namespace ExcelWriter {
  class Chart {
    static readonly AREA_CHART: XLSX.ChartType.AREA_CHART;
    static readonly AREA_STACKED_CHART: XLSX.ChartType.AREA_STACKED_CHART;
    static readonly AREA_STACKED_PERCENT_CHART: XLSX.ChartType.AREA_STACKED_PERCENT_CHART;
    static readonly BAR_CHART: XLSX.ChartType.BAR_CHART;
    static readonly BAR_STACKED_CHART: XLSX.ChartType.BAR_STACKED_CHART;
    static readonly BAR_STACKED_PERCENT_CHART: XLSX.ChartType.BAR_STACKED_PERCENT_CHART;
    static readonly COLUMN_CHART: XLSX.ChartType.COLUMN_CHART;
    static readonly COLUMN_STACKED_CHART: XLSX.ChartType.COLUMN_STACKED_CHART;
    static readonly COLUMN_STACKED_PERCENT_CHART: XLSX.ChartType.COLUMN_STACKED_PERCENT_CHART;
    static readonly DOUGHNUT_CHART: XLSX.ChartType.DOUGHNUT_CHART;
    static readonly LINE_CHART: XLSX.ChartType.LINE_CHART;
    static readonly LINE_STACKED_CHART: XLSX.ChartType.LINE_STACKED_CHART;
    static readonly LINE_STACKED_PERCENT_CHART: XLSX.ChartType.LINE_STACKED_PERCENT_CHART;
    static readonly PIE_CHART: XLSX.ChartType.PIE_CHART;
    static readonly SCATTER_CHART: XLSX.ChartType.SCATTER_CHART;
    static readonly SCATTER_STRAIGHT_CHART: XLSX.ChartType.SCATTER_STRAIGHT_CHART;
    static readonly SCATTER_STRAIGHT_WITH_MARKERS_CHART: XLSX.ChartType.SCATTER_STRAIGHT_WITH_MARKERS_CHART;
    static readonly SCATTER_SMOOTH_CHART: XLSX.ChartType.SCATTER_SMOOTH_CHART;
    static readonly SCATTER_SMOOTH_WITH_MARKERS_CHART: XLSX.ChartType.SCATTER_SMOOTH_WITH_MARKERS_CHART;
    static readonly RADAR_CHART: XLSX.ChartType.RADAR_CHART;
    static readonly RADAR_WITH_MARKERS_CHART: XLSX.ChartType.RADAR_WITH_MARKERS_CHART;
    static readonly RADAR_FILLED_CHART: XLSX.ChartType.RADAR_FILLED_CHART;
    addSeries(categories: string | null, values: string): void;
    setTitleName(name: string): void;
    setTitleNameFont(font: ChartFont): void;
  }

  class Format {
    static readonly NONE_ALIGN: XLSX.Alignment.NONE_ALIGN;
    static readonly LEFT_ALIGN: XLSX.Alignment.LEFT_ALIGN;
    static readonly CENTER_ALIGN: XLSX.Alignment.CENTER_ALIGN;
    static readonly RIGHT_ALIGN: XLSX.Alignment.RIGHT_ALIGN;
    static readonly FILL_ALIGN: XLSX.Alignment.FILL_ALIGN;
    static readonly JUSTIFY_ALIGN: XLSX.Alignment.JUSTIFY_ALIGN;
    static readonly CENTER_ACROSS_ALIGN: XLSX.Alignment.CENTER_ACROSS_ALIGN;
    static readonly DISTRIBUTED_ALIGN: XLSX.Alignment.DISTRIBUTED_ALIGN;
    static readonly VERTICAL_TOP_ALIGN: XLSX.Alignment.VERTICAL_TOP_ALIGN;
    static readonly VERTICAL_BOTTOM_ALIGN: XLSX.Alignment.VERTICAL_BOTTOM_ALIGN;
    static readonly VERTICAL_CENTER_ALIGN: XLSX.Alignment.VERTICAL_CENTER_ALIGN;
    static readonly VERTICAL_JUSTIFY_ALIGN: XLSX.Alignment.VERTICAL_JUSTIFY_ALIGN;
    static readonly VERTICAL_DISTRIBUTED_ALIGN: XLSX.Alignment.VERTICAL_DISTRIBUTED_ALIGN;

    static readonly NONE_BORDER: XLSX.BorderStyle.NONE_BORDER;
    static readonly THIN_BORDER: XLSX.BorderStyle.THIN_BORDER;
    static readonly MEDIUM_BORDER: XLSX.BorderStyle.MEDIUM_BORDER;
    static readonly DASHED_BORDER: XLSX.BorderStyle.DASHED_BORDER;
    static readonly DOTTED_BORDER: XLSX.BorderStyle.DOTTED_BORDER;
    static readonly THICK_BORDER: XLSX.BorderStyle.THICK_BORDER;
    static readonly DOUBLE_BORDER: XLSX.BorderStyle.DOUBLE_BORDER;
    static readonly HAIR_BORDER: XLSX.BorderStyle.HAIR_BORDER;
    static readonly MEDIUM_DASHED_BORDER: XLSX.BorderStyle.MEDIUM_DASHED_BORDER;
    static readonly DASH_DOT_BORDER: XLSX.BorderStyle.DASH_DOT_BORDER;
    static readonly MEDIUM_DASH_DOT_BORDER: XLSX.BorderStyle.MEDIUM_DASH_DOT_BORDER;
    static readonly DASH_DOT_DOT_BORDER: XLSX.BorderStyle.DASH_DOT_DOT_BORDER;
    static readonly MEDIUM_DASH_DOT_DOT_BORDER: XLSX.BorderStyle.MEDIUM_DASH_DOT_DOT_BORDER;
    static readonly SLANT_DASH_DOT_BORDER: XLSX.BorderStyle.SLANT_DASH_DOT_BORDER;

    static readonly SUPERSCRIPT_FONT: XLSX.ScriptStyle.SUPERSCRIPT_FONT;
    static readonly SUBSCRIPT_FONT: XLSX.ScriptStyle.SUBSCRIPT_FONT;

    static readonly NONE_UNDERLINE: XLSX.UnderlineStyle.NONE_UNDERLINE;
    static readonly SINGLE_UNDERLINE: XLSX.UnderlineStyle.SINGLE_UNDERLINE;
    static readonly DOUBLE_UNDERLINE: XLSX.UnderlineStyle.DOUBLE_UNDERLINE;
    static readonly SINGLE_ACCOUNTING_UNDERLINE: XLSX.UnderlineStyle.SINGLE_ACCOUNTING_UNDERLINE;
    static readonly DOUBLE_ACCOUNTING_UNDERLINE: XLSX.UnderlineStyle.DOUBLE_ACCOUNTING_UNDERLINE;

    setAlign(alignment: Alignment): void;
    setRotation(angle: number): void;
    setTextWrap(): void;
    setIndent(level: number): void;
    setShrink(): void;
    setBgColor(color: number): void;
    setFgColor(color: number): void;
    setBorderColor(color: number): void;
    setBottomColor(color: number): void;
    setTopColor(color: number): void;
    setLeftColor(color: number): void;
    setRightColor(color: number): void;
    setFontColor(color: number): void;
    setFontName(name: string): void;
    setFontScript(style: ScriptStyle): void;
    setFontSize(size: number): void;
    setFontStrikeout(): void;
    setBold(): void;
    setItalic(): void;
    setUnderline(style: UnderlineStyle): void;
    setBorder(style: BorderStyle): void;
    setBottom(style: BorderStyle): void;
    setTop(style: BorderStyle): void;
    setLeft(style: BorderStyle): void;
    setRight(style: BorderStyle): void;
    setNumFormat(format: string): void;
  }

  enum Color {
    /** Black */
    BLACK_COLOR = 0x1000000,
    /** Blue */
    BLUE_COLOR = 0x0000ff,
    /** Brown */
    BROWN_COLOR = 0x800000,
    /** Cyan */
    CYAN_COLOR = 0x00ffff,
    /** Gray */
    GRAY_COLOR = 0x808080,
    /** Green */
    GREEN_COLOR = 0x008000,
    /** Lime */
    LIME_COLOR = 0x00ff00,
    /** Magenta */
    MAGENTA_COLOR = 0xff00ff,
    /** Navy */
    NAVY_COLOR = 0x000080,
    /** Orange */
    ORANGE_COLOR = 0xff6600,
    /** Pink */
    PINK_COLOR = 0xff00ff,
    /** Purple */
    PURPLE_COLOR = 0x800080,
    /** Red */
    RED_COLOR = 0xff0000,
    /** Silver */
    SILVER_COLOR = 0xc0c0c0,
    /** White */
    WHITE_COLOR = 0xffffff,
    /** Yellow */
    YELLOW_COLOR = 0xffff00,
  }

  interface ChartFont {
    /** The chart font name, such as "Arial" or "Calibri". */
    name?: string;
    /** The chart font size. The default is 11. */
    size?: number;
    /** The chart font bold property. */
    bold?: boolean;
    /** The chart font italic property. */
    italic?: boolean;
    /** The chart font underline property. */
    underline?: boolean;
    /** The chart font rotation property. Range: -90 to 90, and 270, 271 and 360:
     *
     *  - The angles -90 to 90 are the normal range shown in the Excel user interface.
     *  - The angle 270 gives a stacked (top to bottom) alignment.
     *  - The angle 271 gives a stacked alignment for East Asian fonts.
     *  - The angle 360 gives an explicit angle of 0 to override the y axis default.
     * */
    rotation?: number;
    /** The chart font color. */
    color?: number;
    /** The chart font pitch family property. Rarely required, set to 0. */
    pitchFamily?: number;
    /** The chart font character set property. Rarely required, set to 0. */
    charset?: number;
    /** The chart font baseline property. Rarely required, set to 0. */
    baseline?: number;
  }

  interface DataValidation {
    /** Set the validation type. */
    validate: ValidationType;
    /** Set the validation criteria type to select the data. */
    criteria?: ValidationCriteria;
    /** Controls whether a data validation is not applied to blank data in the cell. */
    ignoreBlank?: boolean;
    /**
     * This parameter is used to toggle on and off the 'Show input message
     * when cell is selected' option in the Excel data validation dialog. When
     * the option is off an input message is not displayed even if it has been
     * set using inputMessage. It is on by default.
     */
    showInput?: boolean;
    /**
     * This parameter is used to toggle on and off the 'Show error alert
     * after invalid data is entered' option in the Excel data validation
     * dialog. When the option is off an error message is not displayed even
     * if it has been set using errorMessage. It is on by default.
     */
    showError?: boolean;
    /**
     * This parameter is used to specify the type of error dialog that is
     * displayed.
     */
    errorType?: ValidationErrorType;
    /**
     * This parameter is used to toggle on and off the 'In-cell dropdown'
     * option in the Excel data validation dialog. When the option is on a
     * dropdown list will be shown for list validations. It is on by default.
     */
    dropdown?: boolean;
    /**
     * This parameter is used to set the limiting value to which the criteria
     * is applied using a whole or decimal number.
     */
    valueNumber?: number;
    /**
     * This parameter is used to set the limiting value to which the criteria
     * is applied using a cell reference. It is valid for any of the
     * `_FORMULA` validation types.
     */
    valueFormula?: string;
    /**
     * This parameter is used to set a list of strings for a drop down list.
     *
     * Note, the string list is restricted by Excel to 255 characters,
     * including comma separators.
     */
    valueList?: string[];
    /**
     * This parameter is used to set the limiting value to which the date or
     * time criteria is applied.
     */
    valueDatetime?: Date;
    /**
     * This parameter is the same as `valueNumber` but for the minimum value
     * when a `BETWEEN` criteria is used.
     */
    minimumNumber?: number;
    /**
     * This parameter is the same as `valueFormula` but for the minimum value
     * when a `BETWEEN` criteria is used.
     */
    minimumFormula?: string;
    /**
     * This parameter is the same as `valueDatetime` but for the minimum value
     * when a `BETWEEN` criteria is used.
     */
    minimumDatetime?: Date;
    /**
     * This parameter is the same as `valueNumber` but for the maximum value
     * when a `BETWEEN` criteria is used.
     */
    maximumNumber?: number;
    /**
     * This parameter is the same as `valueFormula` but for the maximum value
     * when a `BETWEEN` criteria is used.
     */
    maximumFormula?: string;
    /**
     * This parameter is the same as `valueDatetime` but for the maximum value
     * when a `BETWEEN` criteria is used.
     */
    maximumDatetime?: Date;
    /**
     * The inputTitle parameter is used to set the title of the input message
     * that is displayed when a cell is entered. It has no default value and
     * is only displayed if the input message is displayed. See the
     * `inputMessage` parameter below.
     *
     * The maximum title length is 32 characters.
     */
    inputTitle?: string;
    /**
     * The inputMessage parameter is used to set the input message that is
     * displayed when a cell is entered. It has no default value.
     *
     * The message can be split over several lines using newlines. The maximum
     * message length is 255 characters.
     */
    inputMessage?: string;
    /**
     * The errorTitle parameter is used to set the title of the error message
     * that is displayed when the data validation criteria is not met. The
     * default error title is 'Microsoft Excel'. The maximum title length is
     * 32 characters.
     */
    errorTitle?: string;
    /**
     * The errorMessage parameter is used to set the error message that is
     * displayed when a cell is entered. The default error message is "The
     * value you entered is not valid. A user has restricted values that can
     * be entered into the cell".
     *
     * The message can be split over several lines using newlines. The maximum
     * message length is 255 characters.
     */
    errorMessage?: string;
  }

  interface FilterRule {
    criteria: FilterCriteria;
    valueString?: string;
    value?: number;
  }

  interface RowColOptions {
    /** Hide the row/column. */
    hidden?: boolean;
    /** Outline level. */
    level?: number;
    /** Set the outline row as collapsed. */
    collapsed?: boolean;
  }

  class Worksheet {
    static readonly DEFAULT_ROW_HEIGHT: number;

    static readonly EQUAL_TO_FILTER_CRITERIA: XLSX.FilterCriteria.EQUAL_TO_FILTER_CRITERIA;
    static readonly NOT_EQUAL_TO_FILTER_CRITERIA: XLSX.FilterCriteria.NOT_EQUAL_TO_FILTER_CRITERIA;
    static readonly GREATER_THAN_FILTER_CRITERIA: XLSX.FilterCriteria.GREATER_THAN_FILTER_CRITERIA;
    static readonly LESS_THAN_FILTER_CRITERIA: XLSX.FilterCriteria.LESS_THAN_FILTER_CRITERIA;
    static readonly GREATER_THAN_OR_EQUAL_TO_FILTER_CRITERIA: XLSX.FilterCriteria.GREATER_THAN_OR_EQUAL_TO_FILTER_CRITERIA;
    static readonly LESS_THAN_OR_EQUAL_TO_FILTER_CRITERIA: XLSX.FilterCriteria.LESS_THAN_OR_EQUAL_TO_FILTER_CRITERIA;
    static readonly BLANKS_FILTER_CRITERIA: XLSX.FilterCriteria.BLANKS_FILTER_CRITERIA;
    static readonly NON_BLANKS_FILTER_CRITERIA: XLSX.FilterCriteria.NON_BLANKS_FILTER_CRITERIA;

    static readonly AND_FILTER: XLSX.FilterOperator.AND_FILTER;
    static readonly OR_FILTER: XLSX.FilterOperator.OR_FILTER;

    static readonly INTEGER_VALIDATION_TYPE: XLSX.ValidationType.INTEGER_VALIDATION_TYPE;
    static readonly INTEGER_FORMULA_VALIDATION_TYPE: XLSX.ValidationType.INTEGER_FORMULA_VALIDATION_TYPE;
    static readonly DECIMAL_VALIDATION_TYPE: XLSX.ValidationType.DECIMAL_VALIDATION_TYPE;
    static readonly DECIMAL_FORMULA_VALIDATION_TYPE: XLSX.ValidationType.DECIMAL_FORMULA_VALIDATION_TYPE;
    static readonly LIST_VALIDATION_TYPE: XLSX.ValidationType.LIST_VALIDATION_TYPE;
    static readonly LIST_FORMULA_VALIDATION_TYPE: XLSX.ValidationType.LIST_FORMULA_VALIDATION_TYPE;
    static readonly DATE_VALIDATION_TYPE: XLSX.ValidationType.DATE_VALIDATION_TYPE;
    static readonly DATE_FORMULA_VALIDATION_TYPE: XLSX.ValidationType.DATE_FORMULA_VALIDATION_TYPE;
    static readonly DATE_NUMBER_VALIDATION_TYPE: XLSX.ValidationType.DATE_NUMBER_VALIDATION_TYPE;
    static readonly TIME_VALIDATION_TYPE: XLSX.ValidationType.TIME_VALIDATION_TYPE;
    static readonly TIME_FORMULA_VALIDATION_TYPE: XLSX.ValidationType.TIME_FORMULA_VALIDATION_TYPE;
    static readonly TIME_NUMBER_VALIDATION_TYPE: XLSX.ValidationType.TIME_NUMBER_VALIDATION_TYPE;
    static readonly LENGTH_VALIDATION_TYPE: XLSX.ValidationType.LENGTH_VALIDATION_TYPE;
    static readonly LENGTH_FORMULA_VALIDATION_TYPE: XLSX.ValidationType.LENGTH_FORMULA_VALIDATION_TYPE;
    static readonly CUSTOM_FORMULA_VALIDATION_TYPE: XLSX.ValidationType.CUSTOM_FORMULA_VALIDATION_TYPE;
    static readonly ANY_VALIDATION_TYPE: XLSX.ValidationType.ANY_VALIDATION_TYPE;

    static readonly BETWEEN_VALIDATION_CRITERIA: XLSX.ValidationCriteria.BETWEEN_VALIDATION_CRITERIA;
    static readonly NOT_BETWEEN_VALIDATION_CRITERIA: XLSX.ValidationCriteria.NOT_BETWEEN_VALIDATION_CRITERIA;
    static readonly EQUAL_TO_VALIDATION_CRITERIA: XLSX.ValidationCriteria.EQUAL_TO_VALIDATION_CRITERIA;
    static readonly NOT_EQUAL_TO_VALIDATION_CRITERIA: XLSX.ValidationCriteria.NOT_EQUAL_TO_VALIDATION_CRITERIA;
    static readonly GREATER_THAN_VALIDATION_CRITERIA: XLSX.ValidationCriteria.GREATER_THAN_VALIDATION_CRITERIA;
    static readonly LESS_THAN_VALIDATION_CRITERIA: XLSX.ValidationCriteria.LESS_THAN_VALIDATION_CRITERIA;
    static readonly GREATER_THAN_OR_EQUAL_TO_VALIDATION_CRITERIA: XLSX.ValidationCriteria.GREATER_THAN_OR_EQUAL_TO_VALIDATION_CRITERIA;
    static readonly LESS_THAN_OR_EQUAL_TO_VALIDATION_CRITERIA: XLSX.ValidationCriteria.LESS_THAN_OR_EQUAL_TO_VALIDATION_CRITERIA;

    static readonly STOP_VALIDATION_ERROR_TYPE: XLSX.ValidationErrorType.STOP_VALIDATION_ERROR_TYPE;
    static readonly WARNING_VALIDATION_ERROR_TYPE: XLSX.ValidationErrorType.WARNING_VALIDATION_ERROR_TYPE;
    static readonly INFORMATION_VALIDATION_ERROR_TYPE: XLSX.ValidationErrorType.INFORMATION_VALIDATION_ERROR_TYPE;

    autofilter(
      firstRow: number,
      firstColumn: number,
      lastRow: number,
      lastColumn: number
    ): void;
    filterColumn(column: number, rule: FilterRule): void;
    filterColumn(
      column: number,
      rule1: FilterRule,
      rule2: FilterRule,
      operator: FilterOperator
    ): void;
    filterList(column: number, list: string[]): void;
    dataValidationCell(
      row: number,
      column: number,
      validation: DataValidation
    ): void;
    dataValidationRange(
      firstRow: number,
      firstColumn: number,
      lastRow: number,
      lastColumn: number,
      validation: DataValidation
    ): void;
    freezePanes(row: number, column: number): void;
    splitPanes(vertical: number, horizontal: number): void;
    insertChart(row: number, column: number, chart: Chart): void;
    insertImage(row: number, column: number, image: Uint8Array): void;
    mergeRange(
      firstRow: number,
      firstColumn: number,
      lastRow: number,
      lastColumn: number,
      string: string,
      format?: Format
    ): void;
    setColumn(
      firstColumn: number,
      lastColumn: number,
      width: number,
      format?: Format
    ): void;
    setRow(row: number, height: number, format?: Format): void;
    setRow(
      row: number,
      height: number,
      format: Format | null,
      options: RowColOptions
    ): void;
    setFooter(footer: string): void;
    setHeader(header: string): void;
    setSelection(
      firstRow: number,
      firstColumn: number,
      lastRow: number,
      lastColumn: number
    ): void;
    writeBoolean(
      row: number,
      column: number,
      boolean: boolean,
      format?: Format
    ): void;
    writeDatetime(
      row: number,
      column: number,
      date: Date,
      format?: Format
    ): void;
    writeFormula(
      row: number,
      column: number,
      formula: string,
      format?: Format
    ): void;
    writeNumber(
      row: number,
      column: number,
      number: number,
      format?: Format
    ): void;
    writeString(
      row: number,
      column: number,
      string: string,
      format?: Format
    ): void;
    writeURL(row: number, column: number, url: string, format?: Format): void;
  }

  class Workbook {
    addChart(type: ChartType): Chart;
    addFormat(): Format;
    addWorksheet(name: string): Worksheet;
    close(): ArrayBuffer;

    readonly defaultURLFormat: Format;
  }

  /** Convert an Excel `A1` cell string into a `(row, col)` pair. */
  function cell(cell: string): [number, number];

  type Alignment = XLSX.Alignment;
  type BorderStyle = XLSX.BorderStyle;
  type ChartType = XLSX.ChartType;
  type ScriptStyle = XLSX.ScriptStyle;
  type UnderlineStyle = XLSX.UnderlineStyle;
  type FilterCriteria = XLSX.FilterCriteria;
  type FilterOperator = XLSX.FilterOperator;
  type ValidationType = XLSX.ValidationType;
  type ValidationCriteria = XLSX.ValidationCriteria;
  type ValidationErrorType = XLSX.ValidationErrorType;
}

export = ExcelWriter;
