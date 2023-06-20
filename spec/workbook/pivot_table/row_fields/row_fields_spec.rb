# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::RowFields do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/pivot_table/row_fields/row_fields.xlsx')
  row_fields = xlsx.pivot_table_definitions[0].row_fields

  it 'row_fields has count' do
    expect(row_fields.count).to eq(1)
  end

  it 'field has field_index' do
    expect(row_fields[0].field_index).to eq(0)
  end
end
