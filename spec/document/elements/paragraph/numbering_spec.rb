# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Numbering do
  let(:numbering_font) { OoxmlParser::Parser.parse('spec/document/elements/paragraph/numbering/numbering_font_name.docx') }

  it 'numbering.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/numbering/numbering.docx')
    expect(docx.elements.first.numbering).to be_an_instance_of OoxmlParser::NumberingProperties
    expect(docx.elements.first.numbering.abstruct_numbering.level_list.first.text.value).to eq('o')
    expect(docx.elements.first.numbering.abstruct_numbering.level_list.first.numbering_format.value).to eq('bullet')
  end

  it 'numbering_paragraph_indent.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/numbering/numbering_paragraph_indent.docx')
    expect(docx.elements[0].numbering.abstruct_numbering
               .level_list.first.paragraph_properties.indent.hanging_indent).to eq(OoxmlParser::OoxmlSize.new(0.63, :centimeter))
  end

  it 'numbering_font_name.docx' do
    expect(numbering_font.elements[0].numbering.abstruct_numbering
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

  it 'numbering_in_shape.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/numbering/numbering_in_shape.docx')
    expect(docx.elements[0].nonempty_runs.first.alternate_content.office2007_content
               .data.text_box.first.numbering.abstruct_numbering
               .level_list.first.text.value).to eq('Ø')
  end

  it 'numbering_in_table.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/numbering/numbering_in_table.docx')
    expect(docx.element_by_description(location: :canvas, type: :table).first.numbering.abstruct_numbering.level_list.first.text.value).to eq('¨')
  end

  it 'numbering_suffix' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/numbering/numbering_suffix.docx')
    expect(docx.element_by_description.first.numbering.abstruct_numbering.level_list.first.suffix.value).to eq(:space)
  end

  it 'numbering_level_current' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/numbering/numbering_suffix.docx')
    expect(docx.element_by_description.first.numbering.numbering_level_current).to be_a(OoxmlParser::NumberingLevel)
  end

  it 'numbering ilvl is an integer' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/numbering/numbering_suffix.docx')
    expect(docx.element_by_description.first.numbering.ilvl).to eq(0)
  end

  it 'numbering ilvl is an integer non-default' do
    expect(numbering_font.elements[1].numbering.ilvl).to eq(1)
  end
end
