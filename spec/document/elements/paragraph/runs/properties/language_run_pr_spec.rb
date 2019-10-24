# frozen_string_literal: true

require 'spec_helper'

describe 'LanguageRunProperties' do
  it 'language_run_pr_spec' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/properties/language/language_run_pr.docx')
    expect(docx.elements.first.nonempty_runs[1].run_properties.language.value).to eq('en-CA')
  end
end
