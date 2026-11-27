"""Native Excel charts for the management results sheet."""
from design_palette import (
    BLACK, BRIDGE_COLORS, CASH_COLORS, FONT, MONTHLY_COLORS, NUMBER_FONT, ORANGE,
    REVENUE_COLORS, SCENARIO_COLORS, SILVER, WHITE, rgb,
)

MONEY = '#,##0;(#,##0);0'
COLUMN = 51
STACKED = 52
LINE = 4


def _font(font, size=9, bold=False):
    font.Name = FONT
    font.Size = size
    font.Bold = bold
    font.Color = rgb(BLACK)


def finalize_chart_fonts(data):
    """Assign script families to native Excel text runs, including axis labels."""
    from xml.dom import minidom

    drawing = "http://schemas.openxmlformats.org/drawingml/2006/main"
    doc = minidom.parseString(data)
    # The chart Font COM object writes explicit runs for all three scripts
    # but does not expose separate typefaces on the axis tick labels.
    for script, family in (("latin", NUMBER_FONT), ("ea", FONT), ("cs", NUMBER_FONT)):
        for node in doc.getElementsByTagNameNS(drawing, script):
            node.setAttribute("typeface", family)
    result = doc.toxml(encoding="UTF-8")
    doc.unlink()
    return result


def _hide_outline(element):
    # Excel stores the last RGB beside noFill when the outline is hidden.
    line = element.Format.Line
    line.ForeColor.RGB = rgb(BLACK)
    line.Visible = 0


def _fill(element, color):
    fill = element.Format.Fill
    fill.Solid()
    fill.ForeColor.RGB = rgb(color)


def _link(ws, area):
    return f"'{ws.Name}'!{ws.Range(area).Address}"


def _chart(ws, name, source, bounds, kind):
    obj = ws.ChartObjects().Add(*bounds)
    obj.Name = "차트_" + name
    chart = obj.Chart
    chart.ChartType = kind
    while chart.SeriesCollection().Count:
        chart.SeriesCollection(1).Delete()
    area = chart.ChartArea
    for surface in (area, chart.PlotArea):
        _fill(surface, WHITE)
        _hide_outline(surface)
    _font(area.Font)
    chart.HasTitle = True
    chart.ChartTitle.Formula = "=" + _link(ws, source["title_cell"])
    _font(chart.ChartTitle.Font, 10.8, True)
    chart.HasLegend = False
    chart.DisplayBlanksAs = 1
    chart.PlotVisibleOnly = False
    return chart


def _rows(source, column):
    first = source["header"] + 1
    return f"{column}{first}:{column}{first + source['count'] - 1}"


def _series(ws, chart, source, column):
    s = chart.SeriesCollection().NewSeries()
    s.Name = "=" + _link(ws, f"{column}{source['header']}")
    s.XValues = ws.Range(_rows(source, source["category"]))
    s.Values = ws.Range(_rows(source, column))
    return s


def _bar(s, color):
    _fill(s, color)
    _hide_outline(s)
    s.InvertIfNegative = False


def _labels(s, fmt, position=None):
    s.ApplyDataLabels()
    labels = s.DataLabels()
    labels.ShowValue = True
    labels.ShowLegendKey = False
    labels.NumberFormat = fmt
    _font(labels.Font)
    if position is not None:
        labels.Position = position


def _legend(chart, omit=()):
    chart.HasLegend = True
    legend = chart.Legend
    legend.Position = -4107
    _font(legend.Font)
    for entry in sorted(omit, reverse=True):
        legend.LegendEntries(entry).Delete()


def _axes(chart, count, bars, dated=False):
    axis = chart.Axes(1)
    axis.CategoryType = 2
    axis.TickLabelPosition = -4134  # Below the plot even when values cross zero.
    axis.MajorTickMark = -4142
    axis.MinorTickMark = -4142
    axis.TickLabelSpacing = max(1, (count + 5) // 6) if dated else 1
    _hide_outline(axis)
    _font(axis.TickLabels.Font)
    if dated:
        axis.TickLabels.NumberFormat = "yy-mm"
    value = chart.Axes(2)
    value.HasMajorGridlines = not bars
    _hide_outline(value)
    value.MajorTickMark = -4142
    value.MinorTickMark = -4142
    if bars:
        value.TickLabelPosition = -4142
        group = chart.ChartGroups(1)
        group.GapWidth = 55
        group.Overlap = 100
    else:
        _font(value.TickLabels.Font)
        value.TickLabels.NumberFormat = "General"
        grid = value.MajorGridlines.Format.Line
        grid.ForeColor.RGB = rgb(SILVER)
        grid.Weight = 0.5


def _line(s, color, dashed=False):
    s.MarkerStyle = -4142
    s.Format.Line.Visible = -1
    s.Format.Line.ForeColor.RGB = rgb(color)
    s.Format.Line.Weight = 1.75
    if dashed:
        s.Format.Line.DashStyle = 4


def _annual(ws, source, bounds):
    cols = source["values"][:-1]
    chart = _chart(ws, "연도별매출", source, bounds, COLUMN if len(cols) == 1 else STACKED)
    for i, col in enumerate(cols):
        s = _series(ws, chart, source, col)
        _bar(s, REVENUE_COLORS[i % len(REVENUE_COLORS)])
    if len(cols) == 1:
        _labels(s, MONEY, 2)
        plan_label = s.Points(source["count"]).DataLabel
    else:
        total = _series(ws, chart, source, source["values"][-1])
        total.ChartType = LINE
        total.MarkerStyle = -4142
        _hide_outline(total)
        _labels(total, MONEY, 0)  # Above the cumulative stack.
        plan_label = total.Points(source["count"]).DataLabel
        _legend(chart, (len(cols) + 1,))
    plan_label.Font.Bold = True
    plan_label.Font.Color = rgb(ORANGE)
    _axes(chart, source["count"], True)
    chart.Axes(1).TickLabels.NumberFormat = "0"


def _bridge(ws, source, bounds):
    chart = _chart(ws, "브리지", source, bounds, STACKED)
    grouped = source.get("grouped", False)
    for i, col in enumerate(source["values"]):
        s = _series(ws, chart, source, col)
        if i == 0:
            _fill(s, WHITE)
            s.Format.Fill.Visible = 0
            _hide_outline(s)
        else:
            _bar(s, BRIDGE_COLORS[i - 1])
            if i == 1:
                _labels(s, '#,##0;(#,##0);;', -4108)
    # The visible column can be only a pixel tall. Use its top as the
    # anchor so the signed amount remains legible for either direction.
    anchor = _series(ws, chart, source, source["label_height"])
    anchor.ChartType = LINE
    anchor.MarkerStyle = -4142
    _hide_outline(anchor)
    _labels(anchor, MONEY, 0)
    for i in range(1, source["count"] + 1):
        cell = f"{source['label_text']}{source['header'] + i}"
        label = anchor.Points(i).DataLabel
        label.Formula = "=" + _link(ws, cell)
        if grouped and 1 < i < source["count"]:
            # Adjacent changes can be tiny at this scale. Reserve alternate
            # rows above and below the top edge for the signed amounts.
            label.Position = 0 if i % 2 == 0 else 1
            label.Font.Size = 8
    _legend(chart, (1, 5))
    _axes(chart, source["count"], True)
    if grouped:
        ticks = chart.Axes(1).TickLabels
        ticks.Orientation = 60
        ticks.Font.Size = 8
    elif source["count"] > 8:
        chart.Axes(1).TickLabels.Orientation = 45


def _scenario(ws, source, bounds):
    chart = _chart(ws, "DSCR", source, bounds, COLUMN)
    for col, label_col, color in zip(source["values"][:2], source["labels"], SCENARIO_COLORS[:2]):
        s = _series(ws, chart, source, col)
        _bar(s, color)
        _labels(s, MONEY, 2)
        for i in range(1, source["count"] + 1):
            s.Points(i).DataLabel.Formula = "=" + _link(ws, f"{label_col}{source['header'] + i}")
    threshold = _series(ws, chart, source, source["values"][2])
    threshold.ChartType = LINE
    # A markerless line needs two points to draw a segment. Keep the same
    # first source cell in both modes to retain Excel formatting on refresh.
    value_range = _link(ws, _rows(source, source["values"][2]))
    book = ws.Parent
    name = "DashboardBenchmark"
    first = _link(ws, f"{source['values'][2]}{source['header'] + 1}")
    length = f"IF(COUNT({value_range}),ROWS({value_range}),1)"
    book.Names.Add(Name=name, RefersTo=f"={first}:INDEX({value_range},{length})")
    threshold.Values = f"='{book.Name}'!{name}"
    _legend(chart)
    _axes(chart, source["count"], True)
    _line(chart.SeriesCollection(3), SCENARIO_COLORS[2], True)


def _time_series(ws, name, source, bounds, colors):
    chart = _chart(ws, name, source, bounds, LINE)
    for i, (col, color) in enumerate(zip(source["values"], colors)):
        s = _series(ws, chart, source, col)
        _line(s, color, dashed=(name == "현금차입" and i == 2))
    _legend(chart)
    _axes(chart, source["count"], False, dated=True)


def build_charts(ws, m):
    sources = m.chart_sources
    frame = ws.Range("B9:M22")
    gutter = 7
    width = (frame.Width - 3 * gutter) / 4

    def slot(index, span=1):
        return (frame.Left + index * (width + gutter), frame.Top,
                span * width + (span - 1) * gutter, frame.Height)

    _annual(ws, sources["annual"], slot(0))
    has_bridge = sources["bridge"] is not None
    if has_bridge:
        _bridge(ws, sources["bridge"], slot(1))
    _scenario(ws, sources["scenario"], slot(2 if has_bridge else 1))
    _time_series(ws, "월별", sources["monthly"], slot(3 if has_bridge else 2, 1 if has_bridge else 2),
                 MONTHLY_COLORS)
    if "cash" not in sources:
        return
    chart_row, header = m.dashboard_results[-2:]
    bottom = ws.Range(f"B{chart_row}:H{header-2}")
    _time_series(ws, "현금차입", sources["cash"],
                 (bottom.Left, bottom.Top, bottom.Width, bottom.Height), CASH_COLORS)
