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

  it 'cell by nonexisting sheet name' do
    expect(xlsx.cell(0, 0, 'Sheet3')).to be_nil
  end

  it 'sheet name not string nor integer' do
    expect { xlsx.cell(0, 0, xlsx) }.to raise_error(RuntimeError, /Wrong sheet value/)
  end
end
