# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::XlsxHeaderFooter do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/worksheet/header_footer/header_footer.xlsx') }
  let(:header_footer) { xlsx.worksheets.first.header_footer }

  it 'Has align_with_margins' do
    expect(header_footer.align_with_margins).to be_falsey
  end

  it 'Has different_first' do
    expect(header_footer.different_first).to be_truthy
  end

  it 'Has different_odd_even' do
    expect(header_footer.different_odd_even).to be_truthy
  end

  it 'Has scale_with_document' do
    expect(header_footer.scale_with_document).to be_falsey
  end

  it 'Has odd_header' do
    expect(header_footer.odd_header.raw_string).to eq('&L&"Arial"&14Confidential&C&"-,Bold"&Kff0000&D&R&"-,Italic"&UPage &P')
  end

  it 'Has odd_footer' do
    expect(header_footer.odd_footer.raw_string).to eq('&L&S&A&C&YConfidential&R&XPage &P')
  end

  it 'Has even_header' do
    expect(header_footer.even_header.raw_string).to eq('&LConfidential&C&D&RPage &P')
  end

  it 'Has even_footer' do
    expect(header_footer.even_footer.raw_string).to eq('&L&A&CConfidential&RPage &P')
  end

  it 'Has first_header' do
    expect(header_footer.first_header.raw_string).to eq('&L&P&C&N&R&D')
  end

  it 'Has first_footer' do
    expect(header_footer.first_footer.raw_string).to eq('&L&T&C&F&R&A')
  end

  it 'odd_header has left part' do
    expect(header_footer.odd_header.left).to eq('&"Arial"&14Confidential')
  end

  it 'odd_header has center part' do
    expect(header_footer.odd_header.center).to eq('&"-,Bold"&Kff0000&D')
  end

  it 'odd_header has right part' do
    expect(header_footer.odd_header.right).to eq('&"-,Italic"&UPage &P')
  end
end
