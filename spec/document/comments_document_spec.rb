require 'spec_helper'

describe 'CommentsDocument' do
  let(:docx) { OoxmlParser::Parser.parse('spec/document/comments_document/comments_document.docx') }

  it 'have comments' do
    expect(docx.comments_document[0]).to be_a(OoxmlParser::Comment)
  end

  it 'comment have text' do
    expect(docx.comments_document[0].paragraphs.first.nonempty_runs.first.text).to eq('MyComment')
  end
end
