# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PivotTableDefinition do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/pivot_table/pivot_table.xlsx')
  first_table = xlsx.pivot_table_definitions[0]

  it 'pivot table lenght is two' do
    expect(xlsx.pivot_table_definitions.length).to eq(2)
  end

  it 'pivot table is not empty' do
    expect(first_table).to be_a(described_class)
  end

  it 'pivot table name is correct' do
    expect(first_table.name).to eq('PivotTable3')
  end

  it 'pivot table cache id is correct' do
    expect(first_table.cache_id).to eq(0)
  end

  it 'pivot table apply number formats' do
    expect(first_table.apply_number_formats).to be(false)
  end

  it 'pivot table apply border formats' do
    expect(first_table.apply_border_formats).to be(false)
  end

  it 'pivot table apply font format' do
    expect(first_table.apply_font_formats).to be(false)
  end

  it 'pivot table apply pattern format' do
    expect(first_table.apply_pattern_formats).to be(false)
  end

  it 'pivot table apply alignment format' do
    expect(first_table.apply_alignment_formats).to be(false)
  end

  it 'pivot table apply width height formats' do
    expect(first_table.apply_width_height_formats).to be(true)
  end

  it 'pivot table data caption' do
    expect(first_table.data_caption).to eq('Values')
  end

  it 'pivot table use auto formatting' do
    expect(first_table.use_auto_formatting).to be(true)
  end

  it 'pivot table item print titles' do
    expect(first_table.item_print_titles).to be(true)
  end

  it 'pivot table created version' do
    expect(first_table.created_version).to eq(4)
  end

  it 'pivot table indent' do
    expect(first_table.indent).to eq(0)
  end

  it 'pivot table outline' do
    expect(first_table.outline).to be(true)
  end

  it 'pivot table outline data' do
    expect(first_table.outline_data).to be(true)
  end

  it 'pivot table multiple field filters' do
    expect(first_table.multiple_field_filters).to be(false)
  end
end
