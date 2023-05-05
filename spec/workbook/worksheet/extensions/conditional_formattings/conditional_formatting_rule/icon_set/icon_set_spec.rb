# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::IconSet do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions' \
                                   '/conditional_formattings/conditional_formatting_rule/icon_set/icon_set.xlsx')
  icon_set = xlsx.worksheets.first.extension_list[0].conditional_formattings[0].rules[0].icon_set

  it 'Has set' do
    expect(icon_set.set).to eq('5Boxes')
  end

  it 'Has reverse' do
    expect(icon_set.reverse).to be_truthy
  end

  it 'Has show_value' do
    expect(icon_set.show_value).to be_falsey
  end

  it 'Has custom' do
    expect(icon_set.custom).to be_truthy
  end

  it 'Has values' do
    expect(icon_set.values[0]).to be_a(OoxmlParser::ConditionalFormatValueObject)
  end

  it 'Has icons' do
    expect(icon_set.icons[0]).to be_a(OoxmlParser::ConditionalFormattingIcon)
  end

  it 'Value has greater_or_equal' do
    expect(icon_set.values[2].greater_or_equal).to be_falsey
  end
end
