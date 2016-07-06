require 'spec_helper'

describe OoxmlParser::Borders do
  it 'border_size_nil' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/borders/border_size_nil.docx')
    expect(docx.document_styles[39].paragraph_properties.borders.bar).to be_nil
  end

  it 'borders_space_nil' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/borders/borders_space_nil.docx')
    expect(docx.document_styles[25].paragraph_properties.borders.bar.space).to eq(0)
  end

  it 'dropcap_borders.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/borders/dropcap_borders.docx')
    expect(docx.elements[0].borders.top.space.round(1)).to eq(1.2)
    expect(docx.elements[0].borders.right.space.round(1)).to eq(0.8)
    expect(docx.elements[0].borders.bottom.space.round(1)).to eq(2.3)
    expect(docx.elements[0].borders.left.space.round(1)).to eq(0.4)
  end

  it 'border_properties_nil' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/borders/border_properties_nil.docx')
    expect(docx.element_by_description[0].borders.top).to be_nil
    expect(docx.element_by_description[0].borders.left.sz).to eq(1)
    expect(docx.element_by_description[0].borders.right).to be_nil
    expect(docx.element_by_description[0].borders.bottom).to be_nil
  end
end
