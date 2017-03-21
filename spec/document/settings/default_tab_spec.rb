require 'spec_helper'

describe 'DefaultTab' do
  it 'default_tab_stop.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/settings/default_tab/default_tab_stop.docx')
    expect(docx.settings.default_tab_stop).to eq(OoxmlParser::OoxmlSize.new(5, :centimeter))
  end
end
