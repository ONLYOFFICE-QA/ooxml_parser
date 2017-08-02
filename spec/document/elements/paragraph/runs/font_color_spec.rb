require 'spec_helper'

describe 'font_color' do
  it 'color' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/font_color/color.docx')
    expect(docx.elements.first.character_style_array.first.font_color).to eq(OoxmlParser::Color.new(64, 64, 64))
  end

  it 'color_nil' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/font_color/color_nil.docx')
    expect(docx.elements[1].rows.first.cells.first.elements.first.character_style_array.first.font_color).to eq(OoxmlParser::Color.new(nil, nil, nil))
  end

  it 'copied_paragrph_with_color_1' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/font_color/copied_paragrph_with_color_1.docx')
    expect(docx.elements.first.character_style_array.first.font_color).to eq(OoxmlParser::Color.new(197, 216, 240))
    expect(docx.elements.last.character_style_array.first.font_color).to eq(OoxmlParser::Color.new(197, 216, 240))
  end

  it 'copied_paragrph_with_color_2' do
    pending('Color problems. I cannot solve it')
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/font_color/copied_paragrph_with_color_2.docx')
    expect(docx.elements.first.character_style_array.first.font_color).to eq(OoxmlParser::Color.new(191, 191, 191))
    expect(docx.elements.last.character_style_array.first.font_color).to eq(OoxmlParser::Color.new(95, 73, 122))
  end

  it 'Font color 1' do
    # spec/document_editor/one_user/smoke/content/character_smoke_spec.rb
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/font_color/font_color_01.docx')
    docx_color = docx.elements[0].character_style_array[0].font_color
    expect(docx_color).to eq(OoxmlParser::Color.new(197, 216, 240))
  end

  it 'Font color 2' do
    # spec/document_editor/one_user/smoke/content/character_smoke_spec.rb
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/font_color/font_color_02.docx')
    docx_color = docx.elements[0].character_style_array[0].font_color
    expect(docx_color).to eq(OoxmlParser::Color.new(220, 216, 195))
  end

  it 'Font color 3' do
    # spec/document_editor/one_user/smoke/content/character_smoke_spec.rb
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/font_color/font_color_03.docx')
    docx_color = docx.elements[0].character_style_array[0].font_color
    expect(docx_color).to eq(OoxmlParser::Color.new(118, 145, 60))
  end

  it 'Font color 4' do
    # spec/document_editor/one_user/smoke/content/character_smoke_spec.rb
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/font_color/font_color_04.docx')
    docx_color = docx.elements[0].character_style_array[0].font_color
    expect(docx_color).to eq(OoxmlParser::Color.new(54, 97, 145))
  end

  it 'shape_with_font_color' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/font_color/shape_with_font_color.docx')
    expect(docx.elements.first.nonempty_runs.first.alternate_content
               .office2010_content.graphic.data.text_body.elements.first
               .nonempty_runs.first.font_color).to eq(OoxmlParser::Color.new(150, 62, 3))
  end
end
