require 'spec_helper'

describe OoxmlParser::OldDocxShape do
  it 'old_docx_shape' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/alternate_content/office2007_content/old_docx_shape/old_docx_shape.docx')
    expect(docx.elements[29].character_style_array.first.alternate_content.office2007_content.data.elements
               .first.elements.first.elements.first.elements.first.object.file_reference.content.length).to be > 1000
  end

  it 'pict_in_character_run' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/alternate_content/office2007_content/old_docx_shape/pict_in_character_run.docx')
    expect(docx.elements.first.character_style_array[5].shape).not_to be_nil
  end

  it 'picture_target_null' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/alternate_content/office2007_content/old_docx_shape/picture_target_null.docx')
    expect(docx.elements[9].rows.first.cells.first.elements.first.character_style_array[1].drawings.first.graphic.data.path_to_image.path).to be_nil
  end

  it 'no_picture_target' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/alternate_content/office2007_content/old_docx_shape/no_picture_target.docx')
    expect(docx).to be_with_data
  end

  it 'picture_target_no_blip_internals' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/alternate_content/office2007_content/old_docx_shape/picture_target_no_blip_internals.docx')
    expect(docx.elements[13].character_style_array[1].alternate_content.office2010_content.graphic.data.elements[1].object.path_to_image.path).to be_nil
  end
end
