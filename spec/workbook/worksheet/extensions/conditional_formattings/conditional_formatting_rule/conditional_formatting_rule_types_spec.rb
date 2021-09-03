# frozen_string_literal: true

require 'spec_helper'

describe 'Conditional formatting rule types' do
  it 'CellIs has operator' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions/conditional_formattings/conditional_formatting_value_is.xlsx')
    expect(xlsx.worksheets.first.extension_list[0].conditional_formattings[0].rule.operator).to eq(:between)
  end

  it 'CellIs has formulas' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions/conditional_formattings/conditional_formatting_value_is.xlsx')
    expect(xlsx.worksheets.first.extension_list[0].conditional_formattings[0].rule.formulas[0].value).to eq('1')
    expect(xlsx.worksheets.first.extension_list[0].conditional_formattings[0].rule.formulas[1].value).to eq('5')
  end

  it 'Top10 has percent' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions/conditional_formattings/conditional_formatting_top_10.xlsx')
    expect(xlsx.worksheets.first.extension_list[0].conditional_formattings[0].rule.percent).to be_truthy
  end

  it 'Top10 has rank' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions/conditional_formattings/conditional_formatting_top_10.xlsx')
    expect(xlsx.worksheets.first.extension_list[0].conditional_formattings[0].rule.rank).to eq(10)
  end

  it 'Above average has standard deviation' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions/conditional_formattings/conditional_formatting_above_average.xlsx')
    expect(xlsx.worksheets.first.extension_list[0].conditional_formattings[0].rule.standard_deviation).to eq(1)
  end

  it 'Text rule has text' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions/conditional_formattings/conditional_formatting_text.xlsx')
    expect(xlsx.worksheets.first.extension_list[0].conditional_formattings[0].rule.text).to eq('a')
  end
end
