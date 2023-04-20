
external dependencies are set in **projects/smartrename.props**.

For example:
`<ExternalIncludePath>$(VC_IncludePath);$(WindowsSDK_IncludePath);$(CodeLibraries)include;$(ExternalIncludePath)</ExternalIncludePath>`

where `$(CodeLibraries)` is your local repository (eg vcpkg)