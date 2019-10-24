# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'Border Color' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/cell_border_color.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.style.borders.left.color).to eq(OoxmlParser::Color.new(84, 130, 53))
  end

  it 'cell_text_align_top.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/cell_text_align_top.xlsx')
    expect(xlsx.worksheets[0].rows[0].cells[0].style.alignment.vertical).to eq(:top)
  end

  it 'CellFormat' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/cell_format.xlsx')
    expect(xlsx.worksheets[0].rows[0].cells[0].style.numerical_format).to eq('$#,##0.00')
  end

  it 'cell_numerical_general' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/cell_numerical_general.xlsx')
    expect(xlsx.worksheets.first.rows[1].cells[0].style.numerical_format).to eq('General')
  end
end
