require 'spec_helper'

describe OoxmlParser::PageSetup do
  xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/page_setup/simple_page_setup.xlsx')

  it 'PageSetup#paper_size' do
    expect(xlsx.worksheets.first.page_setup.paper_size).to eq(11)
  end

  it 'PageSetup#orientation' do
    expect(xlsx.worksheets.first.page_setup.orientation).to eq(:portrait)
  end
end
