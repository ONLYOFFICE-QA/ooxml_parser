# frozen_string_literal: true

require 'spec_helper'

describe 'Document#content_type' do
  let(:docx) { OoxmlParser::Parser.parse('spec/document/with_data/no_data.docx') }

  it 'content_type has default type' do
    expect(docx.content_types[0]).to be_a(OoxmlParser::ContentTypeDefault)
  end

  it 'content_type has override type' do
    expect(docx.content_types[-1]).to be_a(OoxmlParser::ContentTypeOverride)
  end

  it 'content_type has extension' do
    expect(docx.content_types[0].extension).to eq('bin')
  end

  it 'content_type has content_type' do
    expect(docx.content_types[0].content_type).to eq('application/vnd.openxmlformats-officedocument.oleObject')
  end

  it 'override has partname' do
    expect(docx.content_types[-1].part_name).to eq('/docProps/app.xml')
  end
end
