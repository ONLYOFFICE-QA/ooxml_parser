# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::SheetProtection do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/worksheet/sheet_protection/sheet_protection.xlsx') }

  it 'Has algorithm_name' do
    expect(xlsx.worksheets[0].sheet_protection.algorithm_name).to eq('SHA-512')
  end

  it 'Has hash_value' do
    expect(xlsx.worksheets[0].sheet_protection.hash_value)
      .to eq('ic0GkFgBQ9JNSpXqrsbVpBomduOYBIjRFeFZZ867E7WE6TEuCPZmkPvHIYXYml6FulxWLA1NFCNR3p5TzwaW4g==')
  end

  it 'Has salt_value' do
    expect(xlsx.worksheets[0].sheet_protection.salt_value)
      .to eq('fWjzrxdFYIv7mGfBWqGu2Q==')
  end

  it 'Has spin_count' do
    expect(xlsx.worksheets[0].sheet_protection.spin_count)
      .to eq(100_000)
  end

  it 'Has sheet' do
    expect(xlsx.worksheets[0].sheet_protection.sheet).to be_truthy
  end

  it 'Has auto_filter' do
    expect(xlsx.worksheets[0].sheet_protection.auto_filter).to be_truthy
  end

  it 'Has delete_columns' do
    expect(xlsx.worksheets[0].sheet_protection.delete_columns).to be_truthy
  end

  it 'Has delete_rows' do
    expect(xlsx.worksheets[0].sheet_protection.delete_rows).to be_truthy
  end

  it 'Has format_cells' do
    expect(xlsx.worksheets[0].sheet_protection.format_cells).to be_falsey
  end

  it 'Has format_columns' do
    expect(xlsx.worksheets[0].sheet_protection.format_columns).to be_falsey
  end

  it 'Has format_rows' do
    expect(xlsx.worksheets[0].sheet_protection.format_rows).to be_falsey
  end

  it 'Has insert_columns' do
    expect(xlsx.worksheets[0].sheet_protection.insert_columns).to be_truthy
  end

  it 'Has insert_rows' do
    expect(xlsx.worksheets[0].sheet_protection.insert_rows).to be_truthy
  end

  it 'Has insert_hyperlinks' do
    expect(xlsx.worksheets[0].sheet_protection.insert_hyperlinks).to be_truthy
  end

  it 'Has objects' do
    expect(xlsx.worksheets[0].sheet_protection.objects).to be_truthy
  end

  it 'Has pivot_tables' do
    expect(xlsx.worksheets[0].sheet_protection.pivot_tables).to be_truthy
  end

  it 'Has scenarios' do
    expect(xlsx.worksheets[0].sheet_protection.scenarios).to be_truthy
  end

  it 'Has select_locked_cells' do
    expect(xlsx.worksheets[0].sheet_protection.select_locked_cells).to be_falsey
  end

  it 'Has select_unlocked_cells' do
    expect(xlsx.worksheets[0].sheet_protection.select_unlocked_cells).to be_falsey
  end

  it 'Has sort' do
    expect(xlsx.worksheets[0].sheet_protection.sort).to be_truthy
  end
end
