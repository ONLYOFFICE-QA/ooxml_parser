# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::SheetView do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/sheet_view/sheet_view.xlsx')
  sheet_view = xlsx.worksheets.first.sheet_views.first

  it 'Has top_left_cell' do
    expect(sheet_view.top_left_cell).to eq(OoxmlParser::Coordinates.new(1, 'C'))
  end

  it 'Has workbook_view_id' do
    expect(sheet_view.workbook_view_id).to eq(0)
  end

  it 'Has zoom_scale' do
    expect(sheet_view.zoom_scale).to eq(100)
  end
end
