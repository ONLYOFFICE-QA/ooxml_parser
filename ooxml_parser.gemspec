$LOAD_PATH.unshift File.expand_path('lib', __dir__)
require 'ooxml_parser/version'
Gem::Specification.new do |s|
  s.name = 'ooxml_parser'
  s.version = OoxmlParser::Version::STRING
  s.platform = Gem::Platform::RUBY
  s.required_ruby_version = '>= 2.1'
  s.authors = ['Pavel Lobashov', 'Roman Zagudaev']
  s.summary = 'OoxmlParser Gem'
  s.description = 'Parse OOXML files (docx, xlsx, pptx)'
  s.email = ['shockwavenn@gmail.com', 'rzagudaev@gmail.com']
  s.files = `git ls-files lib LICENSE.txt README.md`.split($RS)
  s.add_runtime_dependency('nokogiri', '~> 1.6')
  s.add_runtime_dependency('ruby-filemagic', '~> 0.1') unless Gem.win_platform?
  s.add_runtime_dependency('rubyzip', '>= 1.1', '< 3.0')
  s.homepage = 'http://rubygems.org/gems/ooxml_parser'
  s.license = 'AGPL-3.0'
end
