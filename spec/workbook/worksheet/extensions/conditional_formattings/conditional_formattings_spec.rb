# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ConditionalFormattings do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions' \
                                   '/conditional_formattings/conditional_formatting_value_is.xlsx')
  conditional_formattings = xlsx.worksheets.first.extension_list[0].conditional_formattings

  it 'Contains array of ConditionalFormatting' do
    expect(conditional_formattings[0]).to be_a(OoxmlParser::ConditionalFormatting)
  end
end
