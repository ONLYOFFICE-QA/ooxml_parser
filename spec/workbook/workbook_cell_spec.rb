# frozen_string_literal: true

require 'spec_helper'

describe 'Workbook#cell' do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/cell/workbook_cell.xlsx') }

  it 'cell by sheet number' do
    expect(xlsx.cell(2, 2, 0).text).to eq('Sheet1')
  end

  it 'cell by sheet name' do
    expect(xlsx.cell(1, 4, 'Sheet2').text).to eq('Sheet2')
  end
end
