# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'ligatures' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/properties/ligatures/ligatures.docx')
    expect(docx.elements.first.nonempty_runs.first.run_properties.ligatures.value).to eq(:all)
  end
end
