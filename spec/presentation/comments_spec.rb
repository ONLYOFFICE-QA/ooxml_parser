require 'spec_helper'

describe 'My behaviour' do
  it 'PresentationComment' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/comments/comments.pptx')
    expect(pptx.comments.first.text).to eq('Is it true?')
  end
end
