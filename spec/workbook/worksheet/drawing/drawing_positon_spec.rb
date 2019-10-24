# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'chart_grouping.xlsx' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/drawing/position/drawing_positon_offste_units.xlsx')
    expect(xlsx.worksheets[0].drawings[0].from.column_offset).to eq(OoxmlParser::OoxmlSize.new(0.47, :centimeter))
  end
end
