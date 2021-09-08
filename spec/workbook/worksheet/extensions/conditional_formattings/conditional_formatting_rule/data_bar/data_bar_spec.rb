# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DataBar do
  let(:xlsx) do
    OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions'\
                              '/conditional_formattings/conditional_formatting_rule/data_bar/data_bar.xlsx')
  end
  let(:data_bar) { xlsx.worksheets.first.extension_list[0].conditional_formattings[0].rules[0].data_bar }

  it 'Has min_length' do
    expect(data_bar.min_length).to eq(0)
  end

  it 'Has max_length' do
    expect(data_bar.max_length).to eq(100)
  end

  it 'Has show_value' do
    expect(data_bar.show_value).to be_falsey
  end

  it 'Has axis_position' do
    expect(data_bar.axis_position).to eq(:middle)
  end

  it 'Has direction' do
    expect(data_bar.direction).to eq(:context)
  end

  it 'Has gradient' do
    expect(data_bar.gradient).to be_falsey
  end

  it 'Has border' do
    expect(data_bar.border).to be_truthy
  end

  it 'Has negative_bar_same_as_positive' do
    expect(data_bar.negative_bar_same_as_positive).to be_truthy
  end

  it 'Has negative_border_same_as_positive' do
    expect(data_bar.negative_border_same_as_positive).to be_falsey
  end

  it 'Has fill_color' do
    expect(data_bar.fill_color).to eq(OoxmlParser::Color.new(99, 142, 198))
  end

  it 'Has negative_fill_color' do
    expect(data_bar.negative_fill_color).to eq(OoxmlParser::Color.new(99, 142, 198))
  end

  it 'Has border_color' do
    expect(data_bar.border_color).to eq(OoxmlParser::Color.new(99, 142, 198))
  end

  it 'Has negative_border_color' do
    expect(data_bar.negative_border_color).to eq(OoxmlParser::Color.new(255, 0, 0))
  end

  it 'Has axis_color' do
    expect(data_bar.axis_color).to eq(OoxmlParser::Color.new(0, 0, 0))
  end

  it 'Has minimum' do
    expect(data_bar.minimum).to be_a(OoxmlParser::ConditionalFormatValueObject)
  end

  it 'Has maximum' do
    expect(data_bar.maximum).to be_a(OoxmlParser::ConditionalFormatValueObject)
  end
end
