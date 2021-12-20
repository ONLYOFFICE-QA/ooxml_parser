# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Selection do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/worksheet/sheet_view/sheet_view.xlsx') }
  let(:selection) { xlsx.worksheets.first.sheet_views.first.selection }

  it 'Has active_cell' do
    expect(selection.active_cell).to eq(OoxmlParser::Coordinates.new(1, 'C'))
  end

  it 'Has active_cell_id' do
    expect(selection.active_cell_id).to eq(0)
  end

  it 'Has reference_sequence' do
    expect(selection.reference_sequence).to eq('C:C')
  end
end
