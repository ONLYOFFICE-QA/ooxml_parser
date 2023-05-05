# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Sheet do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/sheet/sheet.xlsx')

  it 'Has name' do
    expect(xlsx.sheets[0].name).to eq('Sheet1')
  end

  it 'Has sheet_id' do
    expect(xlsx.sheets[0].sheet_id).to eq(1)
  end

  it 'Visible has state' do
    expect(xlsx.sheets[0].state).to eq(:visible)
  end

  it 'Hidden has state' do
    expect(xlsx.sheets[1].state).to eq(:hidden)
  end

  it 'Has id' do
    expect(xlsx.sheets[0].id).to eq('rId1')
  end
end
