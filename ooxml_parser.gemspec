$LOAD_PATH.unshift File.expand_path('../lib', __FILE__)
require 'ooxml_parser/version'
Gem::Specification.new do |s|
  s.name        = 'ooxml_parser'
  s.version     = OoxmlParser::Version::STRING
  s.platform = Gem::Platform::RUBY
  s.required_ruby_version = '>= 1.9'
  s.authors     = ['Pavel Lobashov', 'Roman Zagudaev']
  s.summary     = 'OoxmlParser Gem'
  s.description = 'Parse OOXML files (docx, xlsx, pptx)'
  s.email       = ['shockwavenn@gmail.com', 'rzagudaev@gmail.com']
  s.files = `git ls-files`.split($RS).reject do |file|
    file =~ %r{^(?:
    spec/.*
    |Gemfile
    |Rakefile
    |\.rspec
    |\.gitignore
    |\.rubocop.yml
    |\.rubocop_todo.yml
    |\.travis.yml
    |.*\.eps
    )$}x
  end
  s.add_runtime_dependency('nokogiri', '~> 1.6')
  s.add_runtime_dependency('rubyzip', '~> 1.1')
  s.add_runtime_dependency('xml-simple', '~> 1.1')
  s.homepage    =
      'http://rubygems.org/gems/ooxml_parser'
end