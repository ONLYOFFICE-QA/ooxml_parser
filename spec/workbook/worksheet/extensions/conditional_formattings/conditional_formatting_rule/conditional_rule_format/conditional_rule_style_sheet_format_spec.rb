# frozen_string_literal: true

require 'spec_helper'

describe 'Conditional rule format in style sheet' do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions' \
                                   '/conditional_formattings/conditional_formatting_stop_if_true.xlsx')
  rule = xlsx.worksheets.first.conditional_formattings[0].rules[0]

  it 'Rule contains format_index' do
    expect(rule.format_index).to eq(1)
  end

  it 'Font' do
    expect(rule.format.font.color).to eq(OoxmlParser::Color.new(156, 0, 6))
  end

  it 'Fill background color' do
    expect(rule.format.fill.pattern_fill.background_color).to eq(OoxmlParser::Color.new(255, 199, 206))
  end
end
