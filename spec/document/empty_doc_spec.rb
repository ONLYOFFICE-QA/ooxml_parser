# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'Check Error for Empty zip docx' do
    expect { OoxmlParser::DocxParser.parse_docx('spec/document/other/empty_zip_docx.docx') }.to raise_error(LoadError)
  end
end
