# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PageMargins do
  xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/page_margins/simple_page_margins.xlsx')

  it 'PageMargins#top' do
    expect(xlsx.worksheets.first.page_margins.top.value).to eq(0.75196850393700787)
  end

  it 'PageMargins#top is in inches' do
    expect(xlsx.worksheets.first.page_margins.top.unit).to eq(:inch)
  end

  it 'PageMargins#bottom' do
    expect(xlsx.worksheets.first.page_margins.bottom.value).to eq(0.75196850393700787)
  end

  it 'PageMargins#left' do
    expect(xlsx.worksheets.first.page_margins.left.value).to eq(0.70078740157480324)
  end

  it 'PageMargins#right' do
    expect(xlsx.worksheets.first.page_margins.right.value).to eq(0.70078740157480324)
  end

  it 'PageMargins#header' do
    expect(xlsx.worksheets.first.page_margins.header.value).to eq(0.5)
  end

  it 'PageMargins#footer' do
    expect(xlsx.worksheets.first.page_margins.footer.value).to eq(0.5)
  end
end
