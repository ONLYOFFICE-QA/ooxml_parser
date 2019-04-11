require 'spec_helper'

describe 'Specific Files' do
  it 'Do not crash on gradient stop with unknown SchemeColor' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/specific_files/unknown_scheme_color.xlsx')
    expect(xlsx).to be_with_data
  end
end
