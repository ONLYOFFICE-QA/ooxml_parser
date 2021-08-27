# frozen_string_literal: true

require "#{Dir.pwd}/lib/ooxml_parser/name"
require "#{Dir.pwd}/lib/ooxml_parser/version"

Gem::Specification.new do |s|
  s.name = OoxmlParser::Name::STRING
  s.version = OoxmlParser::Version::STRING
  s.platform = Gem::Platform.new(%w[mingw32])
  s.required_ruby_version = '>= 2.5'
  s.authors = ['ONLYOFFICE', 'Pavel Lobashov', 'Roman Zagudaev']
  s.email = %w[shockwavenn@gmail.com rzagudaev@gmail.com]
  s.summary = 'OoxmlParser Gem'
  s.description = 'Parse OOXML files (docx, xlsx, pptx)'
  s.homepage = "https://github.com/onlyoffice/#{s.name}"
  s.metadata = {
    'bug_tracker_uri' => "#{s.homepage}/issues",
    'changelog_uri' => "#{s.homepage}/blob/master/CHANGELOG.md",
    'documentation_uri' => "https://www.rubydoc.info/gems/#{s.name}",
    'homepage_uri' => s.homepage,
    'source_code_uri' => s.homepage
  }
  s.files = Dir['lib/**/*']
  s.license = 'AGPL-3.0'
  s.add_runtime_dependency('nokogiri', '~> 1')
  s.add_runtime_dependency('rubyzip', '~> 2')
  s.add_development_dependency('overcommit', '~> 0')
  s.add_development_dependency('parallel_tests', '~> 3')
  s.add_development_dependency('rake', '~> 13')
  s.add_development_dependency('rspec', '~> 3')
  s.add_development_dependency('rubocop', '~> 1')
  s.add_development_dependency('rubocop-performance', '~> 1')
  s.add_development_dependency('rubocop-rake', '~> 0')
  s.add_development_dependency('rubocop-rspec', '~> 2')
  s.add_development_dependency('simplecov', '~> 0')
  s.add_development_dependency('yard', '>= 0.9.20')
end
