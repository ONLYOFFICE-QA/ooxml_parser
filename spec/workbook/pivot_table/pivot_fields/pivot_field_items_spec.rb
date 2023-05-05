# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Items do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/pivot_table/pivot_fields.xlsx')
  items = xlsx.pivot_table_definitions[0].pivot_fields[0].items

  it 'there is two item' do
    expect(items.items.length).to eq(2)
  end

  it 'items is not empty' do
    expect(items).to be_a(described_class)
  end

  it 'item has index' do
    expect(items[0].index).to eq(0)
  end

  it 'item has type' do
    expect(items[1].type).to eq(:default)
  end
end
