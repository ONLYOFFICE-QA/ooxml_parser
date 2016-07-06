require 'spec_helper'

describe OoxmlParser::Spacing do
  it 'ShapeParagraphSpacing_1' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/spacing/shape_paragraph_spacing_1.docx')
    expect(docx.element_by_description(location: :canvas, type: :shape).first.spacing).to eq(OoxmlParser::Spacing.new(0.0, 0.35, 1.0, :auto))
  end

  it 'ShapeParagraphSpacing_2' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/spacing/shape_paragraph_spacing_2.docx')
    expect(docx.element_by_description(location: :canvas, type: :shape).first.spacing).to eq(OoxmlParser::Spacing.new(10.0, 0.35, 1.0, :auto))
  end

  it 'ShapeParagrpaphSpacing_3' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/spacing/shape_paragraph_spacing_3.docx')
    expect(docx.element_by_description(location: :canvas, type: :shape).first.spacing).to eq(OoxmlParser::Spacing.new(0, 0.35, 1.5, :auto))
  end

  it 'ShapeDefaultSpacing' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/spacing/space_default_spacing.docx')
    expect(docx.element_by_description(location: :canvas, type: :shape).first.spacing).to eq(OoxmlParser::Spacing.new(0, 0.35, 1.15, :auto))
  end

  it 'TableLineSpacing' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/spacing/table_line_spacing.docx')
    expect(docx.element_by_description(location: :canvas, type: :table)[0].spacing.line).to eq(1.0)
  end

  it 'Parse default spacing' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/spacing/default_spacing.docx')
    expect(docx.elements.first.spacing).to eq(OoxmlParser::Spacing.new(0, 0.35, 1.15, :auto))
  end

  describe 'contextual spacing' do
    it 'contextual_spacing false' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/spacing/contextual_spacing_false.docx')
      expect(docx.elements.first.contextual_spacing).to be_falsey
    end

    it 'contextual_spacing true' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/spacing/contextual_spacing_true.docx')
      expect(docx.document_style_by_name('Title').paragraph_properties.contextual_spacing).to be_truthy
    end
  end
end
