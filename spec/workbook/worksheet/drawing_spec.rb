# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'one_cell_anchor' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/one_cell_anchor.xlsx')
    expect(xlsx.worksheets.first.drawings.first.graphic_frame.graphic_data.first.axises.first.title.elements.first.runs.first.text).to eq('Horizontal Title')
  end
end
