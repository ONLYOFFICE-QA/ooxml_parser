require 'spec_helper'

describe 'My behaviour' do
  it 'Cell Bold Style' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/font/font_bold.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.font_style.bold).to be_truthy
  end

  it 'Font Color 255-255-255' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/font/font_color_255_255_255.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.color).to eq(OoxmlParser::Color.new(255, 255, 255))
  end

  it 'Font Color Theme 237-125-49' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/font/font_color_237_125_49.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.color).to eq(OoxmlParser::Color.new(237, 125, 49))
  end

  it 'Font Color Standart 255-0-0' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/font/font_color_standart_255_0_0.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.color).to eq(OoxmlParser::Color.new(255, 0, 0))
  end

  it 'Font Color Black' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/font/font_color_black.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.color).to eq(OoxmlParser::Color.new(0, 0, 0))
  end

  it 'font_strikeout.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/font/font_strikeout.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.font_style.strike).to eq(:single)
  end

  it 'underline_single' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/font/underline_single.xlsx')
    expect(xlsx.worksheets[0].rows[1].cells[31].style.font.font_style.underlined).to eq(OoxmlParser::Underline.new(:single))
  end

  it 'cell_style_no_font_name' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/font/cell_style_no_font_name.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.font.name).to eq('Calibri')
  end
end
