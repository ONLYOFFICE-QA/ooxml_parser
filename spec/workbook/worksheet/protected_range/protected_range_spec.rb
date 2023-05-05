# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ProtectedRange do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/protected_range/protected_range.xlsx')

  it 'Has algorithm_name' do
    expect(xlsx.worksheets[0].protected_ranges[0].algorithm_name).to eq('SHA-512')
  end

  it 'Has hash_value' do
    expect(xlsx.worksheets[0].protected_ranges[0].hash_value)
      .to eq('4ojY6CSXAyCP7GrcQSG4mZNFUs6tIXCyG/19F0Jhriw+xKvHsNSUUEhlBmZKoWwyQAElSQ500fUQw2JcbLMi7A==')
  end

  it 'Has salt_value' do
    expect(xlsx.worksheets[0].protected_ranges[0].salt_value)
      .to eq('G8t5ny3ARKayZCZMB/aYEg==')
  end

  it 'Has spin_count' do
    expect(xlsx.worksheets[0].protected_ranges[0].spin_count)
      .to eq(100_000)
  end

  it 'Has name' do
    expect(xlsx.worksheets[0].protected_ranges[0].name).to eq('new_range')
  end

  it 'Has reference_sequence' do
    expect(xlsx.worksheets[0].protected_ranges[0].reference_sequence).to eq('$A$1:$B$2')
  end
end
