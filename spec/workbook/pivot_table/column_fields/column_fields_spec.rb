# frozen_string_literal: true

require 'spec_helper'

describe 'column_fields' do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/pivot_table/column_fields/column_fields.xlsx')
  column_fields = xlsx.pivot_table_definitions[0].column_fields

  it 'column_fields has count' do
    expect(column_fields.count).to eq(1)
  end

  it 'field has field_index' do
    expect(column_fields[0].field_index).to eq(0)
  end
end
