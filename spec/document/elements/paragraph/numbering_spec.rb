require 'spec_helper'

describe OoxmlParser::Numbering do
  it 'numbering.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/numbering/numbering.docx')
    expect(docx.elements.first.numbering).to be_an_instance_of OoxmlParser::Numbering
    expect(docx.elements.first.numbering.numbering_properties.ilvls.first.level_text).to eq('o')
    expect(docx.elements.first.numbering.numbering_properties.ilvls.first.num_format).to eq('bullet')
  end
end
