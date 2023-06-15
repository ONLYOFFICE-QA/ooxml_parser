# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DataFields do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/pivot_table/data_fields/data_fields.xlsx')
  data_fields = xlsx.pivot_table_definitions[0].data_fields

  it 'data_fields has count' do
    expect(data_fields.count).to eq(2)
  end

  it 'data_field has base_field' do
    expect(data_fields[0].base_field).to eq(0)
  end

  it 'data_field has base_item' do
    expect(data_fields[0].base_item).to eq(0)
  end

  it 'data_field has field' do
    expect(data_fields[0].field).to eq(6)
  end

  it 'data_field has name' do
    expect(data_fields[0].name).to eq('Sum of Cost')
  end

  it 'data_field has number_format_id' do
    expect(data_fields[0].number_format_id).to eq(10)
  end

  it 'data_field has extension_list' do
    expect(data_fields[0].extension_list[0].data_field.pivot_show_as).to eq(:percentOfParentRow)
  end
end
