# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Protection do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/rows/cells/style/protection/protection.xlsx')
  protection = xlsx.worksheets[0].rows[0].cells[0].style.protection

  it 'Has locked' do
    expect(protection.locked).to be_falsey
  end

  it 'Has hidden' do
    expect(protection.hidden).to be_truthy
  end
end
