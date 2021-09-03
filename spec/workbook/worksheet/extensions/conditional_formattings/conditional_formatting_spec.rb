# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ConditionalFormatting do
  let(:xlsx) do
    OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions'\
                              '/conditional_formattings/conditional_formatting_value_is.xlsx')
  end
  let(:conditional_formatting) { xlsx.worksheets.first.extension_list[0].conditional_formattings[0] }

  it 'Contains rule' do
    expect(conditional_formatting.rule).to be_a(OoxmlParser::ConditionalFormattingRule)
  end

  it 'Contains reference_sequence' do
    expect(conditional_formatting.reference_sequence).to eq('A1')
  end
end
