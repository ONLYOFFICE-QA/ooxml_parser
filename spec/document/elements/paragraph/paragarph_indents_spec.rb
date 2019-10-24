# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Indents do
  it 'indents_round' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/indents/indents_round.docx')
    expect(docx.element_by_description.first.ind.left_indent).to eq(OoxmlParser::OoxmlSize.new(1.25, :centimeter))
    expect(docx.element_by_description.first.ind.right_indent).to eq(OoxmlParser::OoxmlSize.new(0, :centimeter))
  end

  it 'indents_comparasion' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/indents/indents_comparasion.docx')
    expect(docx.element_by_description.first.ind).to eq(OoxmlParser::Indents.new(OoxmlParser::OoxmlSize.new(0, :centimeter),
                                                                                 OoxmlParser::OoxmlSize.new(1.27, :centimeter),
                                                                                 OoxmlParser::OoxmlSize.new(0, :centimeter),
                                                                                 OoxmlParser::OoxmlSize.new(0, :centimeter)))
  end
end
