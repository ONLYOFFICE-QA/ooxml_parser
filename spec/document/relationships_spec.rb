# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'Document Relationship' do
    docx = OoxmlParser::Parser.parse('spec/document/relationships/simple_relationshipt_file.docx')
    expect(docx.relationships[0].target).to eq('styles.xml')
  end
end
