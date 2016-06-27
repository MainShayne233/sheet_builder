# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'sheet_builder/version'

Gem::Specification.new do |spec|
  spec.name          = "sheet_builder"
  spec.version       = SheetBuilder::VERSION
  spec.authors       = ["MainShayne233"]
  spec.email         = ["shaynetremblay@hotmail.com"]

  spec.summary       = 'Sheet Builder is an abstraction built on top of the amazing axlsx gem, and it allows you to '\
                       'to create spreadsheets template, or sheet blueprint in Ruby that are made up of simple arrays '\
                       'of hashes. This you to easily generate spreadsheets on the fly.'

  spec.homepage      = "https://github.com/MainShayne233/sheet_builder"
  spec.license       = "MIT"



  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_development_dependency 'bundler', '~> 1.12'
  spec.add_development_dependency 'rake', '~> 10.0'
  spec.add_development_dependency 'rspec', '~> 3.0'
end
