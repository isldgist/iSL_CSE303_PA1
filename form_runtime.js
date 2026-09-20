(function initializeAnswerIntegrityProtection() {
const protectedSurveyForm = document.getElementById('survey-form');
let clipboardAuditSequence = 0;

async function recordClipboardAttemptInternal(questionNumber, clipboardAction) {
    try {
        const config = await window.surveyConfigPromise;
        const googleScriptUrl = config.form?.googleScriptUrl || '';
        const authToken = document.getElementById('auth-token').value;
        if (!googleScriptUrl || !authToken) {
            return;
        }

        const auditForm = document.createElement('form');
        const auditFrame = document.createElement('iframe');
        const auditFrameName = `clipboard-audit-${Date.now()}-${clipboardAuditSequence++}`;
        auditFrame.name = auditFrameName;
        auditFrame.hidden = true;
        auditForm.method = 'post';
        auditForm.action = googleScriptUrl;
        auditForm.target = auditFrameName;
        auditForm.hidden = true;

        [
            ['action', 'logClipboardAttempt'],
            ['authToken', authToken],
            ['questionNumber', String(questionNumber)],
            ['clipboardAction', clipboardAction]
        ].forEach(function ([name, value]) {
            const input = document.createElement('input');
            input.type = 'hidden';
            input.name = name;
            input.value = value;
            auditForm.appendChild(input);
        });

        document.body.appendChild(auditFrame);
        document.body.appendChild(auditForm);
        HTMLFormElement.prototype.submit.call(auditForm);
        auditForm.remove();
        setTimeout(function () {
            auditFrame.remove();
        }, 30000);
    } catch (error) {
        console.error('복사 및 붙여넣기 시도 기록 실패:', error);
    }
}

const answerIntegrityStates = new WeakMap();
const recentClipboardSignals = new Map();

function getClipboardOperation(clipboardAction) {
    if (clipboardAction === 'copy' || clipboardAction === 'shortcut-c') {
        return 'copy';
    }
    if (clipboardAction === 'paste' || clipboardAction === 'shortcut-v' ||
        clipboardAction.includes('insertFromPaste')) {
        return 'paste';
    }
    if (clipboardAction === 'cut' || clipboardAction === 'shortcut-x' ||
        clipboardAction.includes('deleteByCut')) {
        return 'cut';
    }
    if (clipboardAction === 'drop' || clipboardAction.includes('insertFromDrop')) {
        return 'drop';
    }
    return clipboardAction;
}

function isActualAnswerTransfer(clipboardAction) {
    return [
        'insertFromPaste',
        'insertFromDrop',
        'deleteByCut',
        'untrusted-input',
        'unobserved-value-change',
        'textarea-replaced',
        'integrity-hook-tampered'
    ].includes(clipboardAction);
}

function flagAndRecordClipboardActivityInternal(questionNumber, clipboardAction) {
    const actualTransfer = isActualAnswerTransfer(clipboardAction);
    if (!actualTransfer) {
        return;
    }

    const copyFlag = document.getElementById(`answer-copy-flag-${questionNumber}`);
    if (copyFlag) {
        copyFlag.value = 'true';
    }

    const now = Date.now();
    const signalKey = `${questionNumber}:${getClipboardOperation(clipboardAction)}`;
    const previousSignal = recentClipboardSignals.get(signalKey);
    const isRecent = previousSignal && now - previousSignal.recordedAt < 750;
    if (isRecent && (!actualTransfer || previousSignal.actualTransfer)) {
        return;
    }
    recentClipboardSignals.set(signalKey, {
        recordedAt: now,
        actualTransfer: actualTransfer || Boolean(previousSignal?.actualTransfer)
    });
    recordClipboardAttemptInternal(questionNumber, clipboardAction);
}

function getAnswerDetails(target) {
    if (!(target instanceof HTMLTextAreaElement) ||
        !target.closest('.answer-box')) {
        return null;
    }
    const match = target.id.match(/^written-answer-(\d+)$/);
    if (!match) {
        return null;
    }
    return {
        textarea: target,
        questionNumber: Number(match[1])
    };
}

function recordIntegrityTamper(clipboardAction) {
    document.querySelectorAll('.answer-box textarea').forEach(function (textarea) {
        const answerDetails = getAnswerDetails(textarea);
        if (answerDetails) {
            flagAndRecordClipboardActivityInternal(
                answerDetails.questionNumber,
                clipboardAction
            );
        }
    });
}

function exposeGuardedIntegrityFunction(name, callback) {
    Object.defineProperty(window, name, {
        configurable: false,
        enumerable: false,
        get: function () {
            return callback;
        },
        set: function () {
            recordIntegrityTamper('integrity-hook-tampered');
        }
    });
}

['copy', 'cut', 'paste', 'drop', 'dragstart', 'contextmenu'].forEach(function (eventName) {
    document.addEventListener(eventName, function (event) {
        if (getAnswerDetails(event.target)) {
            event.preventDefault();
        }
    }, true);
});
document.addEventListener('keydown', function (event) {
    const answerDetails = getAnswerDetails(event.target);
    const clipboardShortcut = (event.ctrlKey || event.metaKey) &&
        ['c', 'v', 'x'].includes(event.key.toLowerCase());
    if (answerDetails && clipboardShortcut) {
        event.preventDefault();
    }
}, true);

const syncAnswerIntegrityBaselines = function () {
    document.querySelectorAll('.answer-box textarea').forEach(function (textarea) {
        const state = answerIntegrityStates.get(textarea);
        if (state) {
            state.value = textarea.value;
        }
    });
};

['beforeinput', 'input'].forEach(function (eventName) {
    document.addEventListener(eventName, function (event) {
        const answerDetails = getAnswerDetails(event.target);
        if (!answerDetails) {
            return;
        }

        const clipboardInputTypes = [
            'insertFromPaste',
            'insertFromDrop',
            'deleteByCut'
        ];
        if (clipboardInputTypes.includes(event.inputType)) {
            flagAndRecordClipboardActivityInternal(
                answerDetails.questionNumber,
                eventName === 'input'
                    ? event.inputType
                    : `before-${event.inputType}`
            );
        } else if (!event.isTrusted) {
            flagAndRecordClipboardActivityInternal(
                answerDetails.questionNumber,
                `untrusted-${eventName}`
            );
        }

        if (eventName === 'input') {
            const state = answerIntegrityStates.get(answerDetails.textarea);
            if (state) {
                state.value = answerDetails.textarea.value;
            }
        }
    }, true);
});

document.querySelectorAll('.answer-box').forEach(function (answerBox, index) {
    const writtenAnswer = answerBox.querySelector('textarea');
    const characterCount = answerBox.querySelector('.character-count');
    answerIntegrityStates.set(writtenAnswer, { value: writtenAnswer.value });

    function blockClipboardAction(event) {
        event.preventDefault();
    }

    writtenAnswer.addEventListener('input', function () {
        const maxLength = writtenAnswer.maxLength > 0
            ? writtenAnswer.maxLength
            : 1000;
        characterCount.textContent =
            `${writtenAnswer.value.length} / ${maxLength}자`;
    });
    ['copy', 'cut', 'paste', 'drop', 'dragstart', 'contextmenu'].forEach(function (eventName) {
        writtenAnswer.addEventListener(eventName, function (event) {
            blockClipboardAction(event);
        });
    });
    writtenAnswer.addEventListener('keydown', function (event) {
        const clipboardShortcut = (event.ctrlKey || event.metaKey) &&
            ['c', 'v', 'x'].includes(event.key.toLowerCase());
        if (clipboardShortcut) {
            blockClipboardAction(event);
        }
    });
});

function auditAnswerIntegrity() {
    document.querySelectorAll('.answer-box textarea').forEach(function (textarea) {
        const answerDetails = getAnswerDetails(textarea);
        if (!answerDetails) {
            return;
        }
        const state = answerIntegrityStates.get(textarea);
        if (!state) {
            answerIntegrityStates.set(textarea, { value: textarea.value });
            flagAndRecordClipboardActivityInternal(
                answerDetails.questionNumber,
                'textarea-replaced'
            );
            return;
        }
        if (textarea.value === state.value) {
            return;
        }
        state.value = textarea.value;
        flagAndRecordClipboardActivityInternal(
            answerDetails.questionNumber,
            'unobserved-value-change'
        );

        const characterCount = textarea.parentElement.querySelector('.character-count');
        const maxLength = textarea.maxLength > 0 ? textarea.maxLength : 1000;
        characterCount.textContent = `${textarea.value.length} / ${maxLength}자`;
    });
}

const answerMutationObserver = new MutationObserver(function (mutations) {
    const answerElementChanged = mutations.some(function (mutation) {
        return Array.from(mutation.addedNodes)
            .concat(Array.from(mutation.removedNodes))
            .some(function (node) {
                return node instanceof HTMLTextAreaElement ||
                    (node instanceof Element && node.querySelector('.answer-box textarea'));
            });
    });
    if (answerElementChanged) {
        auditAnswerIntegrity();
    }
});
answerMutationObserver.observe(protectedSurveyForm, { childList: true, subtree: true });

exposeGuardedIntegrityFunction(
    'recordClipboardAttempt',
    recordClipboardAttemptInternal
);
exposeGuardedIntegrityFunction(
    'flagAndRecordClipboardActivity',
    flagAndRecordClipboardActivityInternal
);
exposeGuardedIntegrityFunction(
    'syncAnswerIntegrityBaselines',
    syncAnswerIntegrityBaselines
);
exposeGuardedIntegrityFunction('auditAnswerIntegrity', auditAnswerIntegrity);
setInterval(auditAnswerIntegrity, 500);
})();
