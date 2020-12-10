# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PivotFields do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/pivot_table/pivot_fields/pivot_fields.xlsx') }
  let(:pivot_fields) { xlsx.pivot_table_definitions[0].pivot_fields }

  it 'pivot fields is two' do
    expect(pivot_fields.pivot_field.length).to eq(1)
  end

  it 'pivot fields is not empty' do
    expect(pivot_fields).to be_a(described_class)
  end

  it 'pivot field contains axis info' do
    expect(pivot_fields[0].axis).to eq('axisCol')
  end

  it 'pivot field has show all property' do
    expect(pivot_fields[0].show_all).to eq(false)
  end
end
