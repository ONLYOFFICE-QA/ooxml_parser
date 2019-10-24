# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::OldDocxGroupProperties do
  it 'old_docx_group_wrap_nil' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/alternate_content/office2007_content/data/properties/old_docx_group_wrap_nil.docx')
    expect(docx.elements.first.rows.first.cells.first.elements.first.character_style_array[2]
               .alternate_content.office2007_content.data.properties.wrap).to be_nil
  end
end
