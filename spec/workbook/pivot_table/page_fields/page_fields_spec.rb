# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PageFields do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/pivot_table/page_fields/page_fields.xlsx')
  page_fields = xlsx.pivot_table_definitions[0].page_fields

  it 'page_fields has count' do
    expect(page_fields.count).to eq(2)
  end

  it 'page_field has field' do
    expect(page_fields[0].field).to eq(1)
  end

  it 'page_field has hierarchy' do
    expect(page_fields[0].hierarchy).to eq(-1)
  end
end
