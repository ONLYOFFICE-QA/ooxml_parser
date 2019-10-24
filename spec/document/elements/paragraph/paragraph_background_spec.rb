# frozen_string_literal: true

require 'spec_helper'

describe 'Pagraph Background' do
  it 'inserted_run not empty' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/background/background_color_style_nil.docx')
    expect(docx.elements.first.background_color.style).to eq(:nil)
  end
end
