# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DefinedName do
  xlsx = OoxmlParser::Parser.parse('spec/workbook/defined_name/defined_name.xlsx')

  it 'Has name' do
    expect(xlsx.defined_names[0].name).to eq('Print_Titles')
  end

  it 'Has local_sheet_id' do
    expect(xlsx.defined_names[0].local_sheet_id).to eq(0)
  end

  it 'Has hidden' do
    expect(xlsx.defined_names[0].hidden).to be_falsey
  end

  it 'Has range' do
    expect(xlsx.defined_names[0].range).to eq('Sheet1!$A:$A,Sheet1!$1:$2')
  end
end
