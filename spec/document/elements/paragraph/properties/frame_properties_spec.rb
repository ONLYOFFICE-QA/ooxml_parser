# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::FrameProperties do
  it 'anchor_lock' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/properties/frame_properties/anchor_lock.docx')
    expect(docx.document_styles[17].paragraph_properties.frame_properties.anchor_lock).to be_truthy
  end
end
