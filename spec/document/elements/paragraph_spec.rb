require 'spec_helper'

describe OoxmlParser::DocxParagraph do
  it 'Check Page Break Before' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/page_break_before.docx')
    expect(docx.elements.first.page_break).to eq(true)
  end

  it 'Check Keep Lines Together' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/keep_lines_together.docx')
    expect(docx.elements.first.keep_lines).to eq(true)
  end

  it 'Check Keep Next True' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/keep_next.docx')
    expect(docx.elements.first.keep_next).to eq(true)
  end

  it 'Check Keep Next False' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/keep_lines_together.docx')
    expect(docx.elements.first.keep_next).to eq(false)
  end

  it 'Apply paragraph style for several paragraphs' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/paragraph_style_several_paragraphs.docx')
    docx.elements.last.character_style_array.delete_at(-1) # remove last character style array that always added
    expect(docx.elements.first).to eq(docx.elements.last)
  end

  it 'SectionIndentHangingNilValue' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/section_indent_hanging_nil.docx')
    expect(docx.elements.first.sector_properties.notes.first
               .elements.first.ind).to eq(OoxmlParser::Indents.new(OoxmlParser::OoxmlSize.new(1.02, :centimeter),
                                                                   OoxmlParser::OoxmlSize.new(2.02, :centimeter),
                                                                   OoxmlParser::OoxmlSize.new(3, :centimeter)))
  end

  it 'ParagrphEquals' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/paragraph_equals.docx')
    expect(docx.elements[0]).not_to eq(docx.elements[2])
  end

  it 'table_with_default_table_run_style' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/table_with_default_table_run_style.docx')
    expect(docx.elements.first.rows.first.cells.first.elements.first).not_to be_nil
  end

  it 'instruction_type' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/instruction_type.docx')
    expect(docx.elements.first.nonempty_runs.first.instruction).to eq('MERGEFIELD a')
    expect(docx.element_by_description.first.page_numbering).to eq(false)
  end

  it 'page_numbering' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/page_numbering.docx')
    expect(docx.element_by_description.first.page_numbering).to eq(true)
    expect(docx.element_by_description.first.character_style_array[1].text).to eq('1')
    expect(docx.element_by_description.first.character_style_array[1].page_number).to eq(true)
  end
end
