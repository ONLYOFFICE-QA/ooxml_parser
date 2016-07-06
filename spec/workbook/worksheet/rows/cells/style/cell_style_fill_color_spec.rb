require 'spec_helper'

describe 'My behaviour' do
  it 'Cell Fill Color' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/fill_color/cell_fill_color.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.fill_color.color).to eq(OoxmlParser::Color.new(197, 224, 180))
  end

  it 'fill_color_none.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/fill_color/fill_color_none.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.fill_color.color).to eq(OoxmlParser::Color.new(nil, nil, nil))
  end

  it 'cell_style_fill_color_0.5' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/fill_color/cell_style_fill_color_tint_0.5.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.fill_color.color).to eq(OoxmlParser::Color.new(127, 127, 127))
  end

  it 'cell_style_fill_color_0.35' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/fill_color/cell_style_fill_color_tint_0.35.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.fill_color.color).to eq(OoxmlParser::Color.new(89, 89, 89))
  end

  it 'cell_style_fill_color_0_theme_0' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/fill_color/cell_style_fill_color_tint_0_theme_0.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.fill_color.color).to eq(OoxmlParser::Color.new(255, 255, 255))
  end

  it 'cell_style_fill_color_0_theme_1' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/fill_color/cell_style_fill_color_tint_0_theme_1.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.fill_color.color).to eq(OoxmlParser::Color.new(0, 0, 0))
  end

  it 'fill_color_nil' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/fill_color/fill_color_nil.xlsx')
    expect(xlsx.worksheets.first.rows[10].cells[4].style.fill_color).to be_nil
  end
end
