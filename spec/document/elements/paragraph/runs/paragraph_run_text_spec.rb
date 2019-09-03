require 'spec_helper'

describe 'My behaviour' do
  it 'copy_table_from_cse.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/text/copy_table_from_cse.docx')
    expect(docx.elements.first.rows.first.cells.first.elements.first.character_style_array.first.text).to eq('Simple Test Text')
  end

  it 'nonbreaking_hyphen' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/text/nonbreaking_hyphen.docx')
    expect(docx.elements.first.character_style_array.first.text).to eq('Test–Text')
  end

  it 'tab_inside_run' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/text/tab_inside_run.docx')
    expect(docx.elements[1].character_style_array.first.text).to eq("Run\t\StopRun")
  end

  it 'text_inside_filed' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/text/text_inside_field.docx')
    expect(docx.elements.first.nonempty_runs.first.text).to eq('www.ya.ru')
  end

  it 'Check shape with text' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/text/shape_with_text.docx')
    expect(docx.elements.first.nonempty_runs
               .first.alternate_content.office2007_content
               .data.text_box.text_box_content
               .elements.first
               .character_style_array.first.text)
      .to eq('Lorem ipsum dolor sit amet, consectetur adipiscing elit.'\
             ' Integer consequat faucibus eros, sed mattis tortor consectetur '\
             'cursus. Mauris non eros odio. Curabitur velit metus, placerat sit '\
             'amet tempus cursus, pulvinar sed enim. Vivamus odio arcu, volutpat '\
             'gravida imperdiet vitae, mollis eget augue. Sed ultricies viverra '\
             'convallis. Fusce pharetra mi eget')
  end

  it 'Check ParagraphLineBreak' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/text/text_line_break.docx')
    expect(docx.elements.first.character_style_array.first.text).to eq("Simple Test Text\rSimple Test Text")
  end

  it 'table_merge.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/text/table_merge.docx')
    docx.elements[1].rows[0].cells[0].elements.each do |element|
      expect(element.nonempty_runs.first.text).not_to be_empty
    end
  end

  it 'HeaderFooterSection' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/text/text_in_header_footer_section.docx')
    expect(docx.elements.first.sector_properties.notes.first.elements.first
               .character_style_array.first.text).to eq('Lorem ipsum dolor sit amet, consectetur '\
                              'adipiscing elit. Integer consequat faucibus eros, sed mattis tortor...')
    expect(docx.notes.first.elements.first.character_style_array.first.text).to eq('Simple Test Text')
  end

  it 'text_box_list' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/text/text_box_list.docx')
    expect(docx.elements[160].character_style_array
               .first.shape.text_box.text_box_content
               .elements.first.character_style_array
               .first.text).to include('Guidelines')
  end
end
