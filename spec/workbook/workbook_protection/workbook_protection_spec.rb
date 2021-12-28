# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::WorkbookProtection do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/workbook_protection/workbook_protection.xlsx') }

  it 'Has lock_structure' do
    expect(xlsx.workbook_protection.lock_structure).to be_truthy
  end

  it 'Has workbook_algorithm_name' do
    expect(xlsx.workbook_protection.workbook_algorithm_name).to eq('SHA-512')
  end

  it 'Has workbook_hash_value' do
    expect(xlsx.workbook_protection.workbook_hash_value)
      .to eq('SC0rVhqn9HVfDRfcEFThptSOoKFAi+bXJoOTuU5yGONi2d/nZ6dCl3HUBDp4KKASmzVrUvjZVBZfModJm+aoUw==')
  end

  it 'Has workbook_salt_value' do
    expect(xlsx.workbook_protection.workbook_salt_value)
      .to eq('XurTvgYAAn0dOOmdFAvx0A==')
  end

  it 'Has workbook_spin_count' do
    expect(xlsx.workbook_protection.workbook_spin_count)
      .to eq(100_000)
  end
end
