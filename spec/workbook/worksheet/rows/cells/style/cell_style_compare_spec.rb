# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'cell_style_compare_1' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/style/cell_style_compare/cell_style_compare_1.xlsx')
    rows = xlsx.worksheets.first.rows
    expect(rows[0].cells[0].style).to eq(rows[1].cells[2].style)
  end
end
