require 'spec_helper'

describe 'FontScheme' do
  let(:docx) { OoxmlParser::DocxParser.parse_docx('spec/document/theme/font_scheme/simple_docx.docx') }

  it 'Font scheme name is not empty' do
    expect(docx.theme.font_scheme.name).to eq('Office Classic 2')
  end
end
