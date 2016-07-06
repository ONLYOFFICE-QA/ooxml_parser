require 'spec_helper'

describe 'My behaviour' do
  it 'Добавление новой автофигуры к существующей группе автофигур' do
    # spec/spreadsheet_editor/several_users/collaborative/smoke/shape_collaborative_spec.rb
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/grouping/shape_grouping.xlsx')
    expect(xlsx.worksheets[0].drawings[0].grouping.nil?).not_to eq(true)
    expect(xlsx.worksheets[0].drawings[0].grouping.shapes.length).to eq(1)
    expect(xlsx.worksheets[0].drawings[0].grouping.grouping.shapes.length).to eq(2)
  end

  it 'chart_grouping.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/grouping/chart_grouping.xlsx')
    expect(xlsx.worksheets[0].drawings[0].grouping.nil?).not_to be_nil
    expect(xlsx.worksheets[0].drawings[0].grouping.pictures.length).to eq(2)
  end
end
