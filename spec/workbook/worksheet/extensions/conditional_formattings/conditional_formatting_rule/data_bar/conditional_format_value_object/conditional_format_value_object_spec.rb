# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ConditionalFormatValueObject do
  let(:xlsx) do
    OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions' \
                              '/conditional_formattings/conditional_formatting_rule/data_bar/data_bar.xlsx')
  end
  let(:value) { xlsx.worksheets.first.extension_list[0].conditional_formattings[0].rules[0].data_bar.minimum }

  it 'Has type' do
    expect(value.type).to eq(:num)
  end

  it 'Has formula' do
    expect(value.formula.value).to eq('1')
  end
end
