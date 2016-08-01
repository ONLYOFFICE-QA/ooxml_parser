require 'spec_helper'

describe OoxmlParser::Numbering do
  it 'numbering.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/numbering/numbering.docx')
    expect(docx.elements.first.numbering).to be_an_instance_of OoxmlParser::NumberingProperties
    expect(docx.elements.first.numbering.abstruct_numbering.level_list.first.text.value).to eq('o')
    expect(docx.elements.first.numbering.abstruct_numbering.level_list.first.numbering_format.value).to eq('bullet')
  end

  it 'numbering_paragraph_indent.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/numbering/numbering_paragraph_indent.docx')
    expect(docx.elements[0].numbering.abstruct_numbering
               .level_list.first.paragraph_properties.indent.hanging_indent).to eq(0.63)
  end

  it 'numbering_font_name.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/numbering/numbering_font_name.docx')
    expect(docx.elements[0].numbering.abstruct_numbering
               .level_list.first.run_properties.font_name).to eq('Symbol')
  end

  it 'numbering_multilevel_type.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/numbering/numbering_multilevel_type.docx')
    expect(docx.elements[0].numbering.abstruct_numbering.multilevel_type.value).to eq(:hybrid_multi_level)
  end

  it 'numbering_in_header.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/numbering/numbering_in_header.docx')
    expect(docx.element_by_description(location: :header, type: :paragraph).first.numbering.abstruct_numbering.level_list.first.text.value).to eq('§')
  end
end
