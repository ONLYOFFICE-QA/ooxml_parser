require 'spec_helper'

describe OoxmlParser::Borders do
  it 'border_size_nil' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/borders/border_size_nil.docx')
    expect(docx.document_styles[39].paragraph_properties.paragraph_borders.bar).to be_nil
  end

  it 'borders_space_nil' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/borders/borders_space_nil.docx')
    expect(docx.document_styles[25].paragraph_properties.paragraph_borders.bar.space).to eq(0)
  end

  it 'dropcap_borders.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/borders/dropcap_borders.docx')
    expect(docx.elements[0].borders.top.space).to eq(OoxmlParser::OoxmlSize.new(1.2, :centimeter))
    expect(docx.elements[0].borders.right.space).to eq(OoxmlParser::OoxmlSize.new(0.8, :centimeter))
    expect(docx.elements[0].borders.bottom.space).to eq(OoxmlParser::OoxmlSize.new(2.3, :centimeter))
    expect(docx.elements[0].borders.left.space).to eq(OoxmlParser::OoxmlSize.new(0.4, :centimeter))
  end

  it 'border_properties_nil' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/borders/border_properties_nil.docx')
    expect(docx.element_by_description[0].borders.top).to be_nil
    expect(docx.element_by_description[0].borders.left.size).to eq(OoxmlParser::OoxmlSize.new(1, :point))
    expect(docx.element_by_description[0].borders.right).to be_nil
    expect(docx.element_by_description[0].borders.bottom).to be_nil
  end

  it 'border_size_pt_8' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/borders/border_size_pt_8.docx')
    expect(docx.elements.first.borders.between.sz).to eq(OoxmlParser::OoxmlSize.new(24, :one_eighth_point))
  end
end
