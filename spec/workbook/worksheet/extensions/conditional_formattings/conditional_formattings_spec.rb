# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ConditionalFormattings do
  let(:xlsx) do
    OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions'\
                              '/conditional_formattings/conditional_formatting_value_is.xlsx')
  end
  let(:conditional_formattings) { xlsx.worksheets.first.extension_list[0].conditional_formattings.conditional_formattings_list }

  it 'Contains array of ConditionalFormatting' do
    expect(conditional_formattings[0]).to be_a(OoxmlParser::ConditionalFormatting)
  end
end
