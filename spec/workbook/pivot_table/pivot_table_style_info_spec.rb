# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PivotTableStyleInfo do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/pivot_table/pivot_table.xlsx') }
  let(:style) { xlsx.pivot_table_definitions[0].style_info }

  it 'style info is correct class' do
    expect(style).to be_a(described_class)
  end

  it 'style info has correct name' do
    expect(style.name).to eq('PivotStyleLight16')
  end

  it 'style info show_row_header is correct' do
    expect(style.show_row_header).to be(true)
  end

  it 'style info show_column_header is correct' do
    expect(style.show_column_header).to be(true)
  end

  it 'style info show_row_stripes is correct' do
    expect(style.show_row_stripes).to be(false)
  end

  it 'style info show_column_stripes is correct' do
    expect(style.show_column_stripes).to be(false)
  end

  it 'style info show_last_column is correct' do
    expect(style.show_last_column).to be(true)
  end
end
