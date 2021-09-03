# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ConditionalRuleFormat do
  let(:xlsx) do
    OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions'\
                              '/conditional_formattings/conditional_formatting_value_is.xlsx')
  end
  let(:rule) { xlsx.worksheets.first.extension_list[0].conditional_formattings[0].rule }

  it 'Font' do
    expect(rule.format.font.color).to eq(OoxmlParser::Color.new(156, 0, 6))
  end

  it 'Number format' do
    expect(rule.format.number_format.format_code).to eq('0.00')
  end

  it 'Fill pattern' do
    expect(rule.format.fill.pattern_fill.pattern_type).to eq(:solid)
  end

  it 'Fill foreground color' do
    expect(rule.format.fill.pattern_fill.foreground_color).to eq(OoxmlParser::Color.new(255, 199, 206))
  end

  it 'Fill background color' do
    expect(rule.format.fill.pattern_fill.background_color).to eq(OoxmlParser::Color.new(255, 199, 206))
  end

  it 'Borders style' do
    expect(rule.format.borders.left.style).to eq(:thin)
  end

  it 'Borders color' do
    expect(rule.format.borders.left.color).to eq(OoxmlParser::Color.new(0, 0, 0))
  end
end
