# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'Добавление новой автофигуры к существующей группе автофигур' do
    # spec/spreadsheet_editor/several_users/collaborative/smoke/shape_collaborative_spec.rb
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/grouping/shape_grouping.xlsx')
    expect(xlsx.worksheets[0].drawings[0].grouping.nil?).not_to eq(true)
    expect(xlsx.worksheets[0].drawings[0].grouping.elements.first).to be_a(OoxmlParser::ShapesGrouping)
    expect(xlsx.worksheets[0].drawings[0].grouping.elements.first.elements.length).to eq(2)
  end

  it 'chart_grouping.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/grouping/chart_grouping.xlsx')
    expect(xlsx.worksheets[0].drawings[0].grouping).not_to be_nil
    expect(xlsx.worksheets[0].drawings[0].grouping.elements.length).to eq(2)
  end
end
