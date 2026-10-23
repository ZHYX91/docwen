/* Finder entry point: preserve the frozen onedir layout and forward every argument. */
#include <mach-o/dyld.h>
#include <limits.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <unistd.h>

int main(int argc, char **argv) {
    (void)argc;
    char executable[PATH_MAX], resolved[PATH_MAX], target[PATH_MAX];
    uint32_t size = sizeof(executable);
    if (_NSGetExecutablePath(executable, &size) != 0 || realpath(executable, resolved) == NULL)
        return 126;
    char *slash = strrchr(resolved, '/');
    if (slash == NULL) return 126;
    *slash = '\0';
    int length = snprintf(target, sizeof(target), "%s/../Resources/payload/DocWen", resolved);
    if (length < 0 || (size_t)length >= sizeof(target)) return 126;
    argv[0] = target;
    execv(target, argv);
    perror("Cannot start DocWen");
    return 126;
}
