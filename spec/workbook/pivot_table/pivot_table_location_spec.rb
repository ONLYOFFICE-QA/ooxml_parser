# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PivotTableDefinition do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/pivot_table/pivot_table.xlsx') }
  let(:location) { xlsx.pivot_table_definitions[0].location }

  it 'location is correct class' do
    expect(location).to be_a(OoxmlParser::Location)
  end

  it 'location ref is correct' do
    expect(location.ref).to eq('A3:C20')
  end

  it 'location first_header_row is correct' do
    expect(location.first_header_row).to be(true)
  end

  it 'location first_data_row is correct' do
    expect(location.first_data_row).to be(true)
  end

  it 'location first_data_column is correct' do
    expect(location.first_data_column).to be(false)
  end
end
