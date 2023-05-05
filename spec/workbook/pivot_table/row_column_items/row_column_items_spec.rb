# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ColumnRowItem do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/pivot_table/pivot_fields.xlsx')
  first_table = xlsx.pivot_table_definitions[0]
  row_items = first_table.row_items
  column_items = first_table.column_items

  it 'row items count is 1' do
    expect(row_items.count).to eq(1)
  end

  it 'row items class is correct' do
    expect(row_items[0]).to be_a(described_class)
  end

  it 'column items count is 2' do
    expect(column_items.count).to eq(2)
  end

  it 'column item class is correct' do
    expect(column_items[0]).to be_a(described_class)
  end

  it 'column item type is correct' do
    expect(column_items[1].type).to eq(:grand)
  end
end
