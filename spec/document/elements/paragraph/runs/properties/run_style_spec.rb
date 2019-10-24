# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::RunStyle do
  it 'run_properties_style_name' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/properties/run_style/run_properties_style_name.docx')
    expect(docx.elements.first.nonempty_runs[1].run_properties.run_style.value).to eq('a5')
  end

  it 'run_properties_style_referenced_styles' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/properties/run_style/run_properties_style_referenced_styles.docx')
    expect(docx.elements.first.nonempty_runs[1].run_properties.run_style.referenced.name).to eq('endnote reference')
  end
end
