# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ConditionalFormattingRule do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions' \
                                   '/conditional_formattings/conditional_formatting_value_is.xlsx')
  conditional_formatting = xlsx.worksheets.first.extension_list[0].conditional_formattings[0]

  it 'Has type' do
    expect(conditional_formatting.rules[0].type).to eq(:cellIs)
  end

  it 'Has priority' do
    expect(conditional_formatting.rules[0].priority).to eq(1)
  end

  it 'Has id' do
    expect(conditional_formatting.rules[0].id).to eq('{002C00E3-007C-4E10-AF69-00FE007200BC}')
  end
end
