# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'PresentationComment' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/comments/comments.pptx')
    expect(pptx.comments.list.first.text).to eq('Is it true?')
    expect(pptx.comments.list.first.author.name).to eq('Hamish Mitchell')
  end
end
