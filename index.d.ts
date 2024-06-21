declare namespace XLSX {
  enum Alignment {
    /** No alignment. Cell will use Excel's default for the data type */
    NONE_ALIGN = 0,
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
    NONE_BORDER = 0,
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
    NONE_UNDERLINE = 0,
    /** Single underline */
    SINGLE_UNDERLINE,
    /** Double underline */
    DOUBLE_UNDERLINE,
    /** Single accounting underline */
    SINGLE_ACCOUNTING_UNDERLINE,
    /** Double accounting underline */
    DOUBLE_ACCOUNTING_UNDERLINE,
  }

  class Worksheet {
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

  class Workbook {
    addChart(type: ChartType): Chart;
    addFormat(): Format;
    addWorksheet(name: string): Worksheet;
    close(): ArrayBuffer;

    readonly defaultURLFormat: Format;
  }

  type Alignment = XLSX.Alignment;
  type BorderStyle = XLSX.BorderStyle;
  type ChartType = XLSX.ChartType;
  type ScriptStyle = XLSX.ScriptStyle;
  type UnderlineStyle = XLSX.UnderlineStyle;
  type Worksheet = XLSX.Worksheet;
}

export = ExcelWriter;
