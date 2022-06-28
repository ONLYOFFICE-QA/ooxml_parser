# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ConditionalFormattingIcon do
  let(:xlsx) do
    OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions' \
                              '/conditional_formattings/conditional_formatting_rule/icon_set/icon_set.xlsx')
  end
  let(:icon) { xlsx.worksheets.first.extension_list[0].conditional_formattings[0].rules[0].icon_set.icons[1] }

  it 'Has icon_set' do
    expect(icon.icon_set).to eq('3Symbols2')
  end

  it 'Has icon_id' do
    expect(icon.icon_id).to eq(1)
  end
end
